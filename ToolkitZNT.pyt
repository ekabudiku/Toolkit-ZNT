# -*- #################
import arcpy
import math
import os
import re
import sys

# Deteksi lingkungan kerja (ArcGIS Pro vs ArcMap Desktop) diletakkan di paling atas
IS_PRO = False
try:
    import arcpy.mp as arc_mapping
    IS_PRO = True
except ImportError:
    import arcpy.mapping as arc_mapping
    IS_PRO = False

class Toolbox(object):
    def __init__(self):
        self.label = "Toolkit ZNT (Zona Nilai Tanah)"
        self.alias = "toolkit_znt"
        # Mendaftarkan 4 tools sekaligus
        self.tools = [ZonaTepiJalan, BagiZonaKompak, QCSurveiZona, BatchExportMXD]

# =====================================================================
# TOOL 1: ZONA TEPI JALAN
# =====================================================================
class ZonaTepiJalan(object):
    def __init__(self):
        self.label = "1. Buat Zona Tepi Jalan"
        self.description = "Membuat deliniasi zona tepi jalan: 1 persil pertama di kiri dan kanan jalan dalam radius pencarian."
        self.canRunInBackground = True

    def getParameterInfo(self):
        p0 = arcpy.Parameter(displayName="Jalan (line)", name="road_fc", datatype="GPFeatureLayer", parameterType="Required", direction="Input")
        p0.filter.list = ["Polyline"]
        p1 = arcpy.Parameter(displayName="Persil (polygon)", name="parcel_fc", datatype="GPFeatureLayer", parameterType="Required", direction="Input")
        p1.filter.list = ["Polygon"]
        p2 = arcpy.Parameter(displayName="Output zona tepi jalan", name="out_fc", datatype="DEFeatureClass", parameterType="Required", direction="Output")
        p3 = arcpy.Parameter(displayName="Radius pencarian persil (m)", name="search_radius", datatype="GPDouble", parameterType="Required", direction="Input")
        p3.value = 20.0
        p4 = arcpy.Parameter(displayName="Lebar buffer sisi tanpa persil (m)", name="buffer_dist", datatype="GPDouble", parameterType="Required", direction="Input")
        p4.value = 8.0
        p5 = arcpy.Parameter(displayName="Jarak antar transek / STEP (m)", name="step", datatype="GPDouble", parameterType="Required", direction="Input")
        p5.value = 2.0
        p6 = arcpy.Parameter(displayName="Isi semua lubang", name="fill_all", datatype="GPBoolean", parameterType="Optional", direction="Input")
        p6.value = False
        p7 = arcpy.Parameter(displayName="Luas lubang maks yang diisi (m2)", name="max_hole", datatype="GPDouble", parameterType="Optional", direction="Input")
        p7.value = 500.0
        p8 = arcpy.Parameter(displayName="Sistem koordinat UTM (kosongkan = deteksi otomatis)", name="utm_sr", datatype="GPSpatialReference", parameterType="Optional", direction="Input", category="Lanjutan")
        p9 = arcpy.Parameter(displayName="Transformasi geografis (opsional)", name="geo_transform", datatype="GPString", parameterType="Optional", direction="Input", category="Lanjutan")
        p10 = arcpy.Parameter(displayName="Overlap strip ke persil (m)", name="overlap", datatype="GPDouble", parameterType="Optional", direction="Input", category="Lanjutan")
        p10.value = 0.10
        return [p0, p1, p2, p3, p4, p5, p6, p7, p8, p9, p10]

    def isLicensed(self):
        return True

    def updateParameters(self, parameters):
        parameters[7].enabled = not bool(parameters[6].value)
        return

    def updateMessages(self, parameters):
        checks = [(3, "Radius pencarian"), (4, "Lebar buffer"), (5, "STEP")]
        for i, name in checks:
            v = parameters[i].value
            if v is not None and v <= 0:
                parameters[i].setErrorMessage("%s harus > 0" % name)
        v = parameters[7].value
        if parameters[7].enabled and v is not None and v < 0:
            parameters[7].setErrorMessage("Luas lubang tidak boleh negatif")
        sr = parameters[8].value
        if sr is not None and sr.type != "Projected":
            parameters[8].setErrorMessage("Pilih sistem koordinat terproyeksi (UTM)")
        return

    def execute(self, parameters, messages):
        road_fc = parameters[0].valueAsText
        parcel_fc = parameters[1].valueAsText
        out_fc = parameters[2].valueAsText
        R = float(parameters[3].value)
        BUF = float(parameters[4].value)
        STEP = float(parameters[5].value)
        fill_all = bool(parameters[6].value)
        max_hole = None if fill_all else float(parameters[7].value or 0.0)
        sr_manual = parameters[8].value
        transform = parameters[9].valueAsText or ""
        OVERLAP = float(parameters[10].value) if parameters[10].value is not None else 0.10
        GRID = 50.0

        def msg(t): messages.addMessage(t)

        old_add = arcpy.env.addOutputsToMap
        arcpy.env.addOutputsToMap = False
        arcpy.env.overwriteOutput = True

        if sr_manual is not None:
            SR = sr_manual
            msg("Sistem koordinat manual: %s" % SR.name)
        else:
            SR = self._detect_utm([road_fc, parcel_fc], transform, msg)
        arcpy.env.outputCoordinateSystem = SR
        if transform: arcpy.env.geographicTransformations = transform

        def cells(x0, y0, x1, y1):
            for i in range(int(math.floor(x0 / GRID)), int(math.floor(x1 / GRID)) + 1):
                for j in range(int(math.floor(y0 / GRID)), int(math.floor(y1 / GRID)) + 1):
                    yield (i, j)

        msg("Memuat persil...")
        parcels, pext, grid = {}, {}, {}
        with arcpy.da.SearchCursor(parcel_fc, ["OID@", "SHAPE@"], spatial_reference=SR) as cur:
            for oid, g in cur:
                if g is None or g.area <= 0: continue
                e = g.extent
                parcels[oid] = g
                pext[oid] = (e.XMin, e.YMin, e.XMax, e.YMax)
                for k in cells(e.XMin, e.YMin, e.XMax, e.YMax):
                    grid.setdefault(k, []).append(oid)
        msg("  %d persil dimuat" % len(parcels))

        def cast_ray(P, nx, ny):
            ex, ey = P.X + nx * R, P.Y + ny * R
            bx0, bx1 = min(P.X, ex), max(P.X, ex)
            by0, by1 = min(P.Y, ey), max(P.Y, ey)
            ray = arcpy.Polyline(arcpy.Array([arcpy.Point(P.X, P.Y), arcpy.Point(ex, ey)]), SR)
            origin = arcpy.PointGeometry(arcpy.Point(P.X, P.Y), SR)
            best_d, best_id, seen = None, None, set()
            for k in cells(bx0, by0, bx1, by1):
                for pid in grid.get(k, ()):
                    if pid in seen: continue
                    seen.add(pid)
                    x0, y0, x1, y1 = pext[pid]
                    if x1 < bx0 or x0 > bx1 or y1 < by0 or y0 > by1: continue
                    pg = parcels[pid]
                    if ray.disjoint(pg): continue
                    inter = ray.intersect(pg, 2)
                    if inter is None or inter.length == 0: inter = ray.intersect(pg, 1)
                    if inter is None or inter.pointCount == 0: continue
                    d = origin.distanceTo(inter)
                    if best_d is None or d < best_d:
                        best_d, best_id = d, pid
            return best_d, best_id

        def station_measures(pl):
            L = pl.length
            n = max(1, int(math.ceil(L / STEP)))
            ms = set(round(L * i / n, 3) for i in range(n + 1))
            acc, prev = 0.0, None
            for p in pl.getPart(0):
                if p is None: continue
                if prev is not None:
                    acc += math.hypot(p.X - prev.X, p.Y - prev.Y)
                    ms.add(round(min(acc, L), 3))
                prev = p
            out = []
            for m in sorted(ms):
                if not out or m - out[-1] > 0.01: out.append(m)
            return out

        def left_normal(pl, m, L):
            a = pl.positionAlongLine(max(0.0, m - 0.5)).firstPoint
            b = pl.positionAlongLine(min(L, m + 0.5)).firstPoint
            dx, dy = b.X - a.X, b.Y - a.Y
            ln = math.hypot(dx, dy)
            if ln == 0: return None
            return -dy / ln, dx / ln

        def quad(p1, p2, p3, p4):
            arr = arcpy.Array([arcpy.Point(*p1), arcpy.Point(*p2), arcpy.Point(*p3), arcpy.Point(*p4), arcpy.Point(*p1)])
            return arcpy.Polygon(arr, SR)

        n_roads = int(arcpy.GetCount_management(road_fc).getOutput(0))
        arcpy.SetProgressor("step", "Memproses jalan...", 0, max(1, n_roads), 1)

        selected, strips = set(), []
        n_hit = n_miss = 0
        with arcpy.da.SearchCursor(road_fc, ["OID@", "SHAPE@"], spatial_reference=SR) as cur:
            for idx, (rid, road) in enumerate(cur, 1):
                arcpy.SetProgressorLabel("Jalan %d dari %d" % (idx, n_roads))
                arcpy.SetProgressorPosition()
                if road is None or road.length <= 0: continue
                for ip in range(road.partCount):
                    pl = arcpy.Polyline(road.getPart(ip), SR)
                    L = pl.length
                    if L <= 0: continue
                    base, side_pts = [], {1: [], -1: []}
                    for m in station_measures(pl):
                        nrm = left_normal(pl, m, L)
                        if nrm is None: continue
                        P = pl.positionAlongLine(m).firstPoint
                        base.append((P.X, P.Y))
                        for s in (1, -1):
                            nx, ny = s * nrm[0], s * nrm[1]
                            d, pid = cast_ray(P, nx, ny)
                            if pid is not None:
                                selected.add(pid)
                                reach = d + OVERLAP
                                n_hit += 1
                            else:
                                reach = BUF
                                n_miss += 1
                            side_pts[s].append((P.X + nx * reach, P.Y + ny * reach))
                    for s in (1, -1):
                        pts = side_pts[s]
                        for i in range(len(base) - 1):
                            q = quad(base[i], base[i + 1], pts[i + 1], pts[i])
                            if q is not None and q.area > 0.001: strips.append(q)
        arcpy.ResetProgressor()

        if not selected and not strips:
            arcpy.env.addOutputsToMap = old_add
            raise arcpy.ExecuteError("Tidak ada zona yang terbentuk. Cek data input.")

        arcpy.SetProgressor("default", "Dissolve dan isi lubang...")
        scratch = arcpy.env.scratchGDB
        tmp = os.path.join(scratch, "zona_tmp")
        diss = os.path.join(scratch, "zona_diss")
        for f in (tmp, diss):
            if arcpy.Exists(f): arcpy.Delete_management(f)

        arcpy.CreateFeatureclass_management(scratch, "zona_tmp", "POLYGON", spatial_reference=SR)
        with arcpy.da.InsertCursor(tmp, ["SHAPE@"]) as ic:
            for pid in selected: ic.insertRow([parcels[pid]])
            for g in strips: ic.insertRow([g])

        arcpy.Dissolve_management(tmp, diss, "", "", "MULTI_PART")
        geom = None
        with arcpy.da.SearchCursor(diss, ["SHAPE@"]) as cur:
            for (g,) in cur:
                geom = g if geom is None else geom.union(g)

        geom, n_filled = self._fill_holes(geom, max_hole, SR)

        if arcpy.Exists(out_fc): arcpy.Delete_management(out_fc)
        arcpy.CreateFeatureclass_management(os.path.dirname(out_fc), os.path.basename(out_fc), "POLYGON", spatial_reference=SR)
        arcpy.AddField_management(out_fc, "ZONA", "TEXT", field_length=20)
        arcpy.AddField_management(out_fc, "LUAS_M2", "DOUBLE")
        arcpy.AddField_management(out_fc, "JML_PRSL", "LONG")
        with arcpy.da.InsertCursor(out_fc, ["SHAPE@", "ZONA", "LUAS_M2", "JML_PRSL"]) as ic:
            ic.insertRow([geom, "TEPI_JALAN", round(geom.area, 2), len(selected)])

        for f in (tmp, diss):
            try: arcpy.Delete_management(f)
            except: pass

        arcpy.env.addOutputsToMap = old_add
        arcpy.ResetProgressor()
        msg("Selesai -> %s" % out_fc)
        return

    @staticmethod
    def _detect_utm(fcs, transform, msg):
        wgs84 = arcpy.SpatialReference(4326)
        lons, lats = [], []
        for fc in fcs:
            d = arcpy.Describe(fc)
            sr = d.spatialReference
            e = d.extent
            for x, y in ((e.XMin, e.YMin), (e.XMax, e.YMax), (e.XMin, e.YMax), (e.XMax, e.YMin)):
                pg = arcpy.PointGeometry(arcpy.Point(x, y), sr)
                pg = pg.projectAs(wgs84, transform) if transform else pg.projectAs(wgs84)
                lons.append(pg.firstPoint.X)
                lats.append(pg.firstPoint.Y)
        lon_c = (min(lons) + max(lons)) / 2.0
        lat_c = (min(lats) + max(lats)) / 2.0
        zone = max(1, min(60, int(math.floor((lon_c + 180.0) / 6.0)) + 1))
        epsg = (32700 if lat_c < 0 else 32600) + zone
        sr_utm = arcpy.SpatialReference(epsg)
        msg("Zona UTM terdeteksi otomatis: %s (EPSG:%d)" % (sr_utm.name, epsg))
        return sr_utm

    @staticmethod
    def _fill_holes(poly, max_area, SR):
        def ring_area(pts):
            a = 0.0
            for i in range(len(pts) - 1):
                a += pts[i].X * pts[i + 1].Y - pts[i + 1].X * pts[i].Y
            return abs(a) / 2.0

        rings_out, filled = arcpy.Array(), 0
        for part in poly:
            rings, ring = [], []
            for p in part:
                if p is None:
                    if ring: rings.append(ring)
                    ring = []
                else: ring.append(arcpy.Point(p.X, p.Y))
            if ring: rings.append(ring)
            for k, r in enumerate(rings):
                if k > 0 and (max_area is None or ring_area(r) <= max_area):
                    filled += 1
                    continue
                rings_out.add(arcpy.Array(r))
        return arcpy.Polygon(rings_out, SR), filled


# =====================================================================
# TOOL 2: PEMBAGIAN ZONA SURVEYOR
# =====================================================================
class BagiZonaKompak(object):
    def __init__(self):
        self.label = "2. Bagi Zona (Sama Rata & Mengelompok)"
        self.description = "Membagi poligon zona menggunakan Recursive Spatial Bisection agar jumlah sama rata dan surveyor tidak melompat terlalu jauh."
        self.canRunInBackground = False

    def getParameterInfo(self):
        param0 = arcpy.Parameter(displayName="Input Poligon Zona", name="in_polygon", datatype="DEFeatureClass", parameterType="Required", direction="Input")
        param0.filter.list = ["Polygon"]
        param1 = arcpy.Parameter(displayName="Jumlah Surveyor", name="jml_surveyor", datatype="GPLong", parameterType="Required", direction="Input")
        param1.value = 3
        param2 = arcpy.Parameter(displayName="Output Feature Class", name="out_polygon", datatype="DEFeatureClass", parameterType="Required", direction="Output")
        return [param0, param1, param2]

    def isLicensed(self):
        return True

    def updateParameters(self, parameters):
        return

    def updateMessages(self, parameters):
        if parameters[1].value and parameters[1].value < 2:
            parameters[1].setErrorMessage("Jumlah surveyor minimal 2.")
        return

    def execute(self, parameters, messages):
        in_fc = parameters[0].valueAsText
        jml_surv = int(parameters[1].value)
        out_fc = parameters[2].valueAsText

        arcpy.env.overwriteOutput = True
        
        arcpy.AddMessage("1. Menyiapkan data output...")
        arcpy.CopyFeatures_management(in_fc, out_fc)
        arcpy.AddField_management(out_fc, "Surveyor", "TEXT", field_length=50)

        data_poligon = []
        with arcpy.da.SearchCursor(out_fc, ["OID@", "SHAPE@XY"]) as cursor:
            for row in cursor:
                data_poligon.append({'oid': row[0], 'x': row[1][0], 'y': row[1][1]})

        if not data_poligon:
            arcpy.AddError("Shapefile kosong!")
            return

        def bisection_clustering(data, k):
            if k <= 1: return [data]
            
            min_x = min(d['x'] for d in data)
            max_x = max(d['x'] for d in data)
            min_y = min(d['y'] for d in data)
            max_y = max(d['y'] for d in data)
            
            if (max_x - min_x) > (max_y - min_y):
                data.sort(key=lambda d: d['x'])
            else:
                data.sort(key=lambda d: d['y'])
                
            k1 = k // 2
            k2 = k - k1
            split_idx = int(len(data) * k1 / float(k))
            
            return bisection_clustering(data[:split_idx], k1) + bisection_clustering(data[split_idx:], k2)

        arcpy.AddMessage("2. Menghitung blok klaster secara rekursif...")
        klaster_hasil = bisection_clustering(data_poligon, jml_surv)

        pembagian_dict = {}
        for index, kelompok in enumerate(klaster_hasil):
            nama_surv = "Surveyor-{}".format(index + 1)
            for item in kelompok:
                pembagian_dict[item['oid']] = nama_surv

        arcpy.AddMessage("3. Merekam hasil pembagian ke tabel atribut...")
        with arcpy.da.UpdateCursor(out_fc, ["OID@", "Surveyor"]) as u_cursor:
            for row in u_cursor:
                if row[0] in pembagian_dict:
                    row[1] = pembagian_dict[row[0]]
                    u_cursor.updateRow(row)

        arcpy.AddMessage("Proses selesai! Zona telah terbagi dalam blok yang sama rata.")
        return


# =====================================================================
# TOOL 3: QC SURVEI ZONA
# =====================================================================
class QCSurveiZona(object):
    def __init__(self):
        self.label = "3. QC Hasil Survei (Jumlah Sampel & RSD)"
        self.description = ("Join titik survei dari CSV ke poligon zona, hitung RSD "
                            "dan status QC per zona, serta rekap surveyor.")
        self.canRunInBackground = True

    # -----------------------------------------------------------------
    def getParameterInfo(self):
        p0 = arcpy.Parameter(displayName="File CSV/XLSX titik survei", name="csv_file",
                             datatype="GPTableView", parameterType="Required",
                             direction="Input")

        p1 = arcpy.Parameter(displayName="Kolom X (Longitude)", name="x_field",
                             datatype="Field", parameterType="Required", direction="Input")
        p1.parameterDependencies = [p0.name]

        p2 = arcpy.Parameter(displayName="Kolom Y (Latitude)", name="y_field",
                             datatype="Field", parameterType="Required", direction="Input")
        p2.parameterDependencies = [p0.name]

        p3 = arcpy.Parameter(displayName="Kolom nilai (yang dihitung RSD-nya)",
                             name="nilai_field", datatype="Field",
                             parameterType="Required", direction="Input")
        p3.parameterDependencies = [p0.name]

        p4 = arcpy.Parameter(displayName="Penilai Tanah", name="surveyor_field",
                             datatype="Field", parameterType="Required", direction="Input")
        p4.parameterDependencies = [p0.name]

        p5 = arcpy.Parameter(displayName="Shapefile zona (polygon)", name="shp_zona",
                             datatype="GPFeatureLayer", parameterType="Required",
                             direction="Input")
        p5.filter.list = ["Polygon"]

        p6 = arcpy.Parameter(displayName="NOZN", name="id_zona_field",
                             datatype="Field", parameterType="Required", direction="Input")
        p6.parameterDependencies = [p5.name]

        p7 = arcpy.Parameter(displayName="Output zona hasil QC", name="output_fc",
                             datatype="DEFeatureClass", parameterType="Required",
                             direction="Output")

        p8 = arcpy.Parameter(displayName="Output titik hasil join (opsional)",
                             name="output_titik", datatype="DEFeatureClass",
                             parameterType="Optional", direction="Output")

        p9 = arcpy.Parameter(displayName="Skala Project", name="skala",
                             datatype="GPString", parameterType="Required",
                             direction="Input")
        p9.filter.type = "ValueList"
        p9.filter.list = ["2500", "10000", "25000"]
        p9.value = "10000"

        p10 = arcpy.Parameter(displayName="Jumlah titik minimum per zona",
                              name="min_titik", datatype="GPLong",
                              parameterType="Required", direction="Input",
                              category="Lanjutan")
        p10.value = 3

        p11 = arcpy.Parameter(displayName="Sistem koordinat CSV", name="csv_sr",
                              datatype="GPSpatialReference", parameterType="Optional",
                              direction="Input", category="Lanjutan")

        return [p0, p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11]

    def isLicensed(self):
        return True

    def updateParameters(self, parameters):
        return

    def updateMessages(self, parameters):
        if parameters[10].value is not None and parameters[10].value < 2:
            parameters[10].setErrorMessage("Minimum titik harus >= 2 agar RSD bisa dihitung")
        if parameters[1].value and parameters[2].value and \
                parameters[1].valueAsText == parameters[2].valueAsText:
            parameters[2].setErrorMessage("Kolom Y tidak boleh sama dengan kolom X")
        return

    # -----------------------------------------------------------------
    def execute(self, parameters, messages):
        csv_file = parameters[0].valueAsText
        x_field = parameters[1].valueAsText
        y_field = parameters[2].valueAsText
        nilai_field = parameters[3].valueAsText
        surveyor_field = parameters[4].valueAsText
        shp_zona = parameters[5].valueAsText
        id_zona_field = parameters[6].valueAsText
        output_fc = parameters[7].valueAsText
        output_titik = parameters[8].valueAsText
        skala = parameters[9].valueAsText
        min_titik = int(parameters[10].value)
        csv_sr = parameters[11].value or arcpy.SpatialReference(4326)

        max_rsd = {"2500": 15.0, "25000": 30.0}.get(skala, 25.0)

        arcpy.env.overwriteOutput = True
        scratch = arcpy.env.scratchGDB
        temp_layer = "csv_virtual_layer"
        temp_points = os.path.join(scratch, "qc_temp_points")
        temp_join = os.path.join(scratch, "qc_temp_join")

        def msg(t):
            messages.addMessage(t)

        try:
            # ---------- 1. CSV -> TITIK ----------
            msg("1. Membaca CSV dan mengonversi ke feature class titik...")
            if arcpy.Exists(temp_layer):
                arcpy.Delete_management(temp_layer)
            arcpy.MakeXYEventLayer_management(csv_file, x_field, y_field, temp_layer, csv_sr)

            if arcpy.Exists(temp_points):
                arcpy.Delete_management(temp_points)
            arcpy.CopyFeatures_management(temp_layer, temp_points)
            n_pt = int(arcpy.GetCount_management(temp_points).getOutput(0))
            msg("   %d titik terbaca" % n_pt)
            if n_pt == 0:
                raise arcpy.ExecuteError("Tidak ada titik terbentuk. Cek kolom X/Y pada CSV.")

            # ---------- 2. SPATIAL JOIN ----------
            msg("2. Spatial join titik ke zona...")
            if arcpy.Exists(temp_join):
                arcpy.Delete_management(temp_join)
            arcpy.SpatialJoin_analysis(target_features=temp_points,
                                       join_features=shp_zona,
                                       out_feature_class=temp_join,
                                       join_operation="JOIN_ONE_TO_MANY",
                                       join_type="KEEP_ALL",
                                       match_option="INTERSECT")

            if output_titik:
                msg("   Menyimpan titik hasil join...")
                arcpy.CopyFeatures_management(temp_join, output_titik)
                arcpy.SetParameterAsText(8, output_titik)

            # ---------- 3. STATISTIK ----------
            msg("3. Menghitung RSD dan merekap surveyor...")
            actual = [f.name for f in arcpy.ListFields(temp_join)]
            col_id = cari_nama_kolom(actual, id_zona_field)
            col_nilai = cari_nama_kolom(actual, nilai_field)
            col_surv = cari_nama_kolom(actual, surveyor_field)

            zona_data = {}
            n_luar = 0
            with arcpy.da.SearchCursor(temp_join, [col_id, col_nilai, col_surv]) as cur:
                for id_z, val, surv in cur:
                    if id_z is None:
                        n_luar += 1
                        continue
                    d = zona_data.setdefault(id_z, {"nilai": [], "surveyor": set()})
                    v = bersihkan_angka(val)
                    if v is not None:
                        d["nilai"].append(v)
                    if surv is not None and unicode(surv).strip() != u"":
                        s = unicode(surv).strip()
                        d["surveyor"].add(s)
            if n_luar:
                messages.addWarningMessage("   %d titik berada di luar semua zona" % n_luar)

            # ---------- 4. OUTPUT ZONA ----------
            msg("4. Menyiapkan output zona (skala %s, batas RSD %.0f%%)..." % (skala, max_rsd))
            arcpy.CopyFeatures_management(shp_zona, output_fc)
            out_fields = [f.name for f in arcpy.ListFields(output_fc)]
            col_id_out = cari_nama_kolom(out_fields, id_zona_field)

            new_fields = [("Jml_Titik", "SHORT", None), ("Rata_Nilai", "DOUBLE", None),
                          ("RSD_Persen", "DOUBLE", None), ("Status_QC", "TEXT", 100),
                          ("Surveyor", "TEXT", 255)]
            existing = [f.name.upper() for f in arcpy.ListFields(output_fc)]
            for name, ftype, flen in new_fields:
                if name.upper() in existing:
                    continue
                if flen:
                    arcpy.AddField_management(output_fc, name, ftype, field_length=flen)
                else:
                    arcpy.AddField_management(output_fc, name, ftype)

            n_valid = n_revisi = n_kurang = n_kosong = 0
            upd = [col_id_out, "Jml_Titik", "Rata_Nilai", "RSD_Persen", "Status_QC", "Surveyor"]
            with arcpy.da.UpdateCursor(output_fc, upd) as cur:
                for row in cur:
                    d = zona_data.get(row[0])
                    if not d:
                        cur.updateRow([row[0], 0, 0.0, 0.0, "Belum Disurvey", "-"])
                        n_kosong += 1
                        continue

                    vals = d["nilai"]
                    n = len(vals)
                    rata2 = sum(vals) / n if n else 0.0
                    rsd = 0.0
                    
                    if n > 1 and rata2 > 0:
                        var = sum((x - rata2) ** 2 for x in vals) / (n - 1)
                        # Fix perhitungan: Tambahkan operator perkalian (*) sebelum 100
                        rsd = (math.sqrt(var) / rata2) * 100.0

                    if n == 0:
                        status = "Belum Disurvey"
                        n_kosong += 1
                    elif n < min_titik:
                        status = "Kurang Lengkap (Tunggu min %d titik)" % min_titik
                        n_kurang += 1
                    elif rsd == 0.0:
                        status = "Perlu Revisi (RSD 0% tidak diizinkan)"
                        n_revisi += 1
                    elif rsd > max_rsd:
                        status = "Perlu Revisi (RSD %.1f%% > batas %d%%)" % (rsd, int(max_rsd))
                        n_revisi += 1
                    else:
                        status = "Lengkap & Valid"
                        n_valid += 1

                    surv_str = ", ".join(sorted(list(d["surveyor"])))
                    if not surv_str:
                        surv_str = "-"
                    surv_str = surv_str[:255] # Limit text length

                    cur.updateRow([row[0], n, rata2, rsd, status, surv_str])

            msg("Rekap zona: %d valid | %d perlu revisi | %d kurang lengkap | %d belum disurvey"
                % (n_valid, n_revisi, n_kurang, n_kosong))
            arcpy.SetParameterAsText(7, output_fc)
            msg("Proses selesai.")

        finally:
            for f in (temp_layer, temp_points, temp_join):
                try:
                    if arcpy.Exists(f):
                        arcpy.Delete_management(f)
                except Exception:
                    pass
        return


# =====================================================================
# FUNGSI UMUM
# =====================================================================
def cari_nama_kolom(list_kolom_aktual, target):
    """Cocokkan nama kolom walau berubah setelah join / dipotong 10 karakter."""
    if not target:
        return ""
    t_low = target.strip().lower()
    t_san = re.sub(r'[^a-zA-Z0-9_]', '_', target).strip().lower()
    for f in list_kolom_aktual:
        if f.lower() == t_low:
            return f
    for f in list_kolom_aktual:
        if f.lower() == t_san:
            return f
    for f in list_kolom_aktual:
        if f.lower() == t_san[:10]:
            return f
    for f in list_kolom_aktual:
        if f.lower().startswith(t_san[:8]):
            return f
    raise ValueError("Kolom '%s' tidak ditemukan. Yang tersedia: %s"
                     % (target, ", ".join(list_kolom_aktual)))


def bersihkan_angka(val):
    """Ubah teks/angka dari format excel (termasuk ribuan ID/EN) menjadi float secara aman."""
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return float(val)
    
    # Deteksi Python 2 'long' integer secara aman
    try:
        if isinstance(val, long):
            return float(val)
    except NameError:
        pass
        
    # Ekstraksi dan sanitasi string (Hindari error "module" atau "type")
    s = unicode(val).strip()
    if not s or s.lower() in ['none', 'null', 'nan']:
        return None
        
    # Buang semua karakter aneh, sisakan angka, koma, titik, minus
    s = re.sub(r'[^\d\,\.\-]', '', s)
    if not s or s == '-':
        return None
        
    # Resolusi format ribuan/desimal
    if s.count('.') >= 1 and s.count(',') == 1:
        s = s.replace('.', '').replace(',', '.')
    elif s.count(',') >= 1 and s.count('.') == 1:
        s = s.replace(',', '')
    elif s.count('.') > 1 and s.count(',') == 0:
        s = s.replace('.', '')
    elif s.count(',') == 1 and s.count('.') == 0:
        s = s.replace(',', '.')
        
    try:
        return float(s)
    except ValueError:
        return None

# =====================================================================
# TOOL 4: BATCH EXPORT MXD
# =====================================================================
class BatchExportMXD(object):
    def __init__(self):
        self.label = "4. Batch Export MXD"
        self.description = "Melakukan export MXD & APRX secara masal dengan opsi format file beserta DPI. Efektif efisien."
        self.canRunInBackground = False

    def getParameterInfo(self):
        param_in_folder = arcpy.Parameter(displayName="Folder Sumber Dokumen Peta", name="in_folder", datatype="DEFolder", parameterType="Required", direction="Input")
        
        param_recursive = arcpy.Parameter(displayName="Pindai Seluruh Subfolder di Dalamnya (Rekursif)?", name="recursive", datatype="GPBoolean", parameterType="Optional", direction="Input")
        param_recursive.value = False

        param_formats = arcpy.Parameter(displayName="Format Ekspor", name="export_formats", datatype="GPString", parameterType="Required", direction="Input", multiValue=True)
        param_formats.filter.type = "ValueList"
        param_formats.filter.list = ["PDF", "JPEG", "PNG", "TIFF"]
        param_formats.value = ["PDF", "JPEG"]

        param_dpi = arcpy.Parameter(displayName="Resolusi Ekspor (DPI)", name="dpi_setting", datatype="GPLong", parameterType="Required", direction="Input")
        param_dpi.value = 300

        param_out_folder = arcpy.Parameter(displayName="Folder Tujuan Output", name="out_folder", datatype="DEFolder", parameterType="Required", direction="Input")

        param_keep_structure = arcpy.Parameter(displayName="Pertahankan Struktur Subfolder pada Output?", name="keep_subfolders", datatype="GPBoolean", parameterType="Optional", direction="Input")
        param_keep_structure.value = True

        return [param_in_folder, param_recursive, param_formats, param_dpi, param_out_folder, param_keep_structure]

    def isLicensed(self):
        return True

    def updateParameters(self, parameters):
        return

    def updateMessages(self, parameters):
        if parameters[3].value is not None and parameters[3].value < 72:
            parameters[3].setErrorMessage("Nilai DPI minimal adalah 72.")
        return

    def _safe_makedirs(self, folder_path):
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)

    def _export_desktop(self, mxd_path, out_base_path, formats, dpi):
        mxd = arc_mapping.MapDocument(mxd_path)
        try:
            for fmt in formats:
                out_file = "{}.{}".format(out_base_path, fmt.lower())
                if fmt == "PDF":
                    arc_mapping.ExportToPDF(mxd, out_file, resolution=dpi)
                elif fmt == "JPEG":
                    arc_mapping.ExportToJPEG(mxd, out_file, resolution=dpi)
                elif fmt == "PNG":
                    arc_mapping.ExportToPNG(mxd, out_file, resolution=dpi)
                elif fmt == "TIFF":
                    arc_mapping.ExportToTIFF(mxd, out_file, resolution=dpi)
        finally:
            del mxd

    def _export_pro(self, doc_path, out_base_path, formats, dpi, scratch_dir):
        ext = os.path.splitext(doc_path)[1].lower()
        aprx = None
        temp_aprx_created = False
        try:
            if ext == ".aprx":
                aprx = arc_mapping.ArcGISProject(doc_path)
            elif ext == ".mxd":
                temp_aprx_path = os.path.join(scratch_dir, "temp_import.aprx")
                if os.path.exists(temp_aprx_path):
                    os.remove(temp_aprx_path)
                try:
                    current_aprx = arc_mapping.ArcGISProject("CURRENT")
                    current_aprx.saveACopy(temp_aprx_path)
                    aprx = arc_mapping.ArcGISProject(temp_aprx_path)
                    temp_aprx_created = True
                except:
                    default_template = os.path.join(arcpy.GetInstallInfo()["InstallDir"], "Resources", "NewProjectTemplates", "Blank.aptx")
                    aprx = arc_mapping.ArcGISProject(default_template)
                aprx.importDocument(doc_path)

            layouts = aprx.listLayouts()
            if not layouts:
                arcpy.AddWarning("  -> Dilewati: Tidak ditemukan Layout pada {}".format(os.path.basename(doc_path)))
                return

            for lyt in layouts:
                suffix = "_{}".format(lyt.name) if len(layouts) > 1 else ""
                for fmt in formats:
                    out_file = "{}{}.{}".format(out_base_path, suffix, fmt.lower())
                    if fmt == "PDF":
                        lyt.exportToPDF(out_file, resolution=dpi)
                    elif fmt == "JPEG":
                        lyt.exportToJPEG(out_file, resolution=dpi)
                    elif fmt == "PNG":
                        lyt.exportToPNG(out_file, resolution=dpi)
                    elif fmt == "TIFF":
                        lyt.exportToTIFF(out_file, resolution=dpi)
        finally:
            if aprx:
                del aprx
            if temp_aprx_created and os.path.exists(temp_aprx_path):
                try:
                    os.remove(temp_aprx_path)
                except:
                    pass

    def execute(self, parameters, messages):
        in_folder = parameters[0].valueAsText
        recursive = bool(parameters[1].value)
        formats = parameters[2].valueAsText.replace("'", "").replace('"', "").split(";")
        dpi = int(parameters[3].value)
        out_folder = parameters[4].valueAsText
        keep_structure = bool(parameters[5].value)

        scratch_dir = arcpy.env.scratchFolder
        engine_name = "ArcGIS Pro (arcpy.mp)" if IS_PRO else "ArcGIS Desktop (arcpy.mapping)"
        arcpy.AddMessage("Menjalankan ekspor menggunakan engine: {}".format(engine_name))

        valid_exts = [".mxd", ".aprx"] if IS_PRO else [".mxd"]

        doc_files = []
        if recursive:
            for root, dirs, files in os.walk(in_folder):
                for f in files:
                    if os.path.splitext(f)[1].lower() in valid_exts:
                        doc_files.append(os.path.join(root, f))
        else:
            for f in os.listdir(in_folder):
                full_path = os.path.join(in_folder, f)
                if os.path.isfile(full_path) and os.path.splitext(f)[1].lower() in valid_exts:
                    doc_files.append(full_path)

        total_files = len(doc_files)
        if total_files == 0:
            arcpy.AddWarning("Tidak ditemukan file dokumen peta ({}) pada folder yang dipilih.".format(", ".join(valid_exts)))
            return

        arcpy.AddMessage("Ditemukan {} file untuk diekspor ke format: {} (DPI: {})".format(total_files, ", ".join(formats), dpi))
        arcpy.SetProgressor("step", "Mengekspor dokumen peta...", 0, total_files, 1)

        success_count = 0
        failed_count = 0

        for idx, doc_path in enumerate(doc_files, 1):
            file_name = os.path.basename(doc_path)
            base_name = os.path.splitext(file_name)[0]
            arcpy.SetProgressorLabel("Memproses ({}/{}): {}".format(idx, total_files, file_name))
            arcpy.AddMessage("[{}/{}] Memproses: {}".format(idx, total_files, file_name))

            if keep_structure and recursive:
                rel_dir = os.path.relpath(os.path.dirname(doc_path), in_folder)
                target_dir = os.path.join(out_folder, rel_dir)
            else:
                target_dir = out_folder

            self._safe_makedirs(target_dir)
            out_base = os.path.join(target_dir, base_name)

            try:
                if IS_PRO:
                    self._export_pro(doc_path, out_base, formats, dpi, scratch_dir)
                else:
                    self._export_desktop(doc_path, out_base, formats, dpi)
                success_count += 1
            except Exception as e:
                arcpy.AddError("  -> Gagal mengekspor {}: {}".format(file_name, str(e)))
                failed_count += 1

            arcpy.SetProgressorPosition()

        arcpy.ResetProgressor()
        arcpy.AddMessage("Selesai! Berhasil: {} file | Gagal: {} file.".format(success_count, failed_count))
