# -*- coding: utf-8 -*-
import arcpy
import csv
import math
import numbers
import os
import re
import sys
import traceback

# --------------------------
# Kompatibilitas tipe string dan angka (Python 2.7 & Python 3 / ArcGIS Pro)
if sys.version_info[0] >= 3:
    unicode = str
    basestring = str
    long = int
    text_type = str
else:
    text_type = unicode

TIDAK_TERIDENTIFIKASI = "Tidak Teridentifikasi"
# --------------------------

# Deteksi lingkungan kerja (ArcGIS Pro vs ArcMap Desktop)
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
        # Daftar 5 tools
        self.tools = [ZonaTepiJalan, BagiZonaKompak, QCSurveiZona, BatchExportMXD, EkstrakZonaExcel]


# =====================================================================
# FUNGSI PENCOCOKAN KOLOM OTOMATIS (AUTO-DETECT FIELDS)
# =====================================================================
def _norm_nama_kolom(s):
    """Normalisasi string nama kolom (hapus simbol/spasi, huruf kecil semua)."""
    if not s:
        return ""
    return re.sub(r'[^a-z0-9]', '', unicode(s).lower())


def tebak_kolom_terbaik(kolom_list, exact_candidates, keyword_candidates=None):
    """
    Mencari nama kolom paling cocok dari daftar kolom yang ada:
    1. Cek kesamaan persis (case-insensitive)
    2. Cek kesamaan setelah simbol/spasi dibersihkan (misal: 'Nilai Tanah/m2' == 'Nilai_Tanah_m2')
    3. Cek apakah mengandung kata kunci (substring)
    """
    if not kolom_list:
        return None

    # 1. Exact match (case-insensitive, trim spasi)
    for cand in exact_candidates:
        c_up = unicode(cand).strip().upper()
        for kol in kolom_list:
            if unicode(kol).strip().upper() == c_up:
                return kol

    # 2. Normalized match (abaikan spasi, garis miring, underscore, titik)
    for cand in exact_candidates:
        c_norm = _norm_nama_kolom(cand)
        if not c_norm:
            continue
        for kol in kolom_list:
            if _norm_nama_kolom(kol) == c_norm:
                return kol

    # 3. Substring / Keyword match
    if keyword_candidates:
        for kw in keyword_candidates:
            kw_norm = _norm_nama_kolom(kw)
            if not kw_norm:
                continue
            for kol in kolom_list:
                if kw_norm in _norm_nama_kolom(kol):
                    return kol

    return None


# =====================================================================
# FUNGSI EKSPOR EXCEL ANTI-GAGAL (DIGUNAKAN OLEH TOOL 3 & TOOL 5)
# =====================================================================
def normalisasi_path_output_excel(raw_path, fallback_ref_path=None):
    """
    Memastikan path output Excel:
    1. Tidak berada di dalam Geodatabase (.gdb / .mdb / .sde)
    2. Memiliki folder tujuan yang valid
    3. Memiliki ekstensi .xlsx atau .xls yang sesuai
    """
    if not raw_path:
        return None

    clean_path = unicode(raw_path).strip().strip("'\"")
    if not clean_path:
        return None

    norm_p = os.path.normpath(clean_path)
    parts = norm_p.split(os.sep)
    gdb_idx = None
    for i, part in enumerate(parts):
        if part.lower().endswith((".gdb", ".mdb", ".sde")):
            gdb_idx = i
            break

    filename = os.path.basename(norm_p)
    if gdb_idx is not None:
        parent_dir = os.sep.join(parts[:gdb_idx])
        if not parent_dir or not os.path.isdir(parent_dir):
            if fallback_ref_path:
                parent_dir = os.path.dirname(os.path.normpath(fallback_ref_path))
                if parent_dir.lower().endswith((".gdb", ".mdb", ".sde")):
                    parent_dir = os.path.dirname(parent_dir)
            else:
                parent_dir = os.path.expanduser("~")
        clean_path = os.path.join(parent_dir, filename)

    out_dir = os.path.dirname(clean_path)
    if not out_dir:
        if fallback_ref_path:
            ref_dir = os.path.dirname(os.path.normpath(fallback_ref_path))
            if ref_dir.lower().endswith((".gdb", ".mdb", ".sde")):
                ref_dir = os.path.dirname(ref_dir)
            out_dir = ref_dir
        else:
            out_dir = arcpy.env.scratchFolder or os.path.expanduser("~")
        clean_path = os.path.join(out_dir, filename)

    if out_dir and not os.path.exists(out_dir):
        try:
            os.makedirs(out_dir)
        except Exception:
            pass

    if not clean_path.lower().endswith((".xlsx", ".xls")):
        has_openpyxl = False
        try:
            import openpyxl
            has_openpyxl = True
        except ImportError:
            pass

        if IS_PRO or has_openpyxl:
            clean_path += ".xlsx"
        else:
            clean_path += ".xls"

    return clean_path


def ekspor_data_ke_excel(headers, rows, field_specs, out_excel_raw, fallback_ref_path=None, messages=None):
    """
    Mengekspor data ke file Excel dengan 3 metode berlapis:
    1. Native Python (openpyxl untuk .xlsx / xlwt untuk .xls)
    2. arcpy.TableToExcel_conversion melalui scratchGDB (otomatis coba .xlsx dan .xls)
    3. Fallback CSV (jika semua metode di atas gagal)
    """
    def log_msg(t):
        if messages:
            messages.addMessage(t)
        else:
            arcpy.AddMessage(t)

    def log_warn(t):
        if messages:
            messages.addWarningMessage(t)
        else:
            arcpy.AddWarning(t)

    out_excel = normalisasi_path_output_excel(out_excel_raw, fallback_ref_path)
    if not out_excel:
        return None

    if os.path.exists(out_excel):
        try:
            os.remove(out_excel)
        except Exception as e_rm:
            base, ext = os.path.splitext(out_excel)
            alt_path = "%s_new%s" % (base, ext)
            log_warn("   File '%s' sedang terbuka/terkunci (%s). Menyimpan ke: %s" % (os.path.basename(out_excel), str(e_rm), alt_path))
            out_excel = alt_path

    ext = os.path.splitext(out_excel)[1].lower()

    # METODE 1A: Gunakan openpyxl jika target .xlsx
    if ext == ".xlsx":
        try:
            import openpyxl
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Ringkasan_QC"
            ws.append([unicode(h) for h in headers])
            for r in rows:
                clean_r = []
                for val in r:
                    if val is None:
                        clean_r.append("")
                    elif isinstance(val, (int, long, float)):
                        clean_r.append(val)
                    else:
                        clean_r.append(unicode(val))
                ws.append(clean_r)
            wb.save(out_excel)
            wb.close()
            log_msg("   Ringkasan Excel (.xlsx) berhasil disimpan ke: %s" % out_excel)
            return out_excel
        except Exception:
            pass

    # METODE 1B: Gunakan xlwt jika target .xls (atau di ArcMap Python 2.7)
    try:
        import xlwt
        xls_path = out_excel if ext == ".xls" else os.path.splitext(out_excel)[0] + ".xls"
        if ext == ".xls" or not IS_PRO:
            if os.path.exists(xls_path):
                try:
                    os.remove(xls_path)
                except Exception:
                    xls_path = os.path.splitext(xls_path)[0] + "_new.xls"
            wb = xlwt.Workbook(encoding="utf-8")
            ws = wb.add_sheet("Ringkasan_QC")
            for c_idx, h in enumerate(headers):
                ws.write(0, c_idx, unicode(h))
            for r_idx, r in enumerate(rows, start=1):
                for c_idx, val in enumerate(r):
                    if val is None:
                        ws.write(r_idx, c_idx, u"")
                    elif isinstance(val, (int, long, float)):
                        ws.write(r_idx, c_idx, val)
                    else:
                        ws.write(r_idx, c_idx, unicode(val)[:32000])
            wb.save(xls_path)
            log_msg("   Ringkasan Excel (.xls) berhasil disimpan ke: %s" % xls_path)
            return xls_path
    except Exception:
        pass

    # METODE 2: Gunakan arcpy.TableToExcel_conversion via scratchGDB
    scratch_gdb = arcpy.env.scratchGDB
    if scratch_gdb and arcpy.Exists(scratch_gdb):
        tbl_name = "tmp_excel_export"
        temp_tbl = os.path.join(scratch_gdb, tbl_name)
        try:
            if arcpy.Exists(temp_tbl):
                arcpy.Delete_management(temp_tbl)
            arcpy.CreateTable_management(scratch_gdb, tbl_name)

            field_names = []
            for fname, ftype, flen in field_specs:
                field_names.append(fname)
                if ftype.upper() == "TEXT":
                    arcpy.AddField_management(temp_tbl, fname, "TEXT", field_length=min(flen or 254, 254))
                else:
                    arcpy.AddField_management(temp_tbl, fname, ftype)

            with arcpy.da.InsertCursor(temp_tbl, field_names) as icur:
                for r in rows:
                    clean_r = []
                    for idx, val in enumerate(r):
                        ftype = field_specs[idx][1].upper()
                        flen = min(field_specs[idx][2] or 254, 254)
                        if ftype == "TEXT":
                            clean_r.append((unicode(val) if val is not None else u"-")[:flen])
                        elif ftype in ("LONG", "SHORT"):
                            clean_r.append(int(val) if val is not None else 0)
                        else:
                            clean_r.append(float(val) if val is not None else 0.0)
                    icur.insertRow(clean_r)

            candidates = [out_excel]
            alt_ext = ".xls" if out_excel.lower().endswith(".xlsx") else ".xlsx"
            candidates.append(os.path.splitext(out_excel)[0] + alt_ext)

            for cand_path in candidates:
                try:
                    if os.path.exists(cand_path):
                        os.remove(cand_path)
                    arcpy.TableToExcel_conversion(temp_tbl, cand_path)
                    if os.path.exists(cand_path):
                        log_msg("   Ringkasan Excel berhasil disimpan ke: %s" % cand_path)
                        return cand_path
                except Exception:
                    continue
        except Exception as e_tte:
            log_warn("   TableToExcel gagal (%s), mencoba fallback CSV..." % str(e_tte))
        finally:
            try:
                if arcpy.Exists(temp_tbl):
                    arcpy.Delete_management(temp_tbl)
            except Exception:
                pass

    # METODE 3: Fallback ke CSV
    try:
        csv_path = os.path.splitext(out_excel)[0] + ".csv"
        if sys.version_info[0] >= 3:
            with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                writer.writerow([unicode(h) for h in headers])
                for r in rows:
                    writer.writerow([unicode(v) if v is not None else "" for v in r])
        else:
            with open(csv_path, "wb") as f:
                f.write(b"\xef\xbb\xbf")
                writer = csv.writer(f)
                writer.writerow([unicode(h).encode("utf-8") for h in headers])
                for r in rows:
                    writer.writerow([(unicode(v) if v is not None else u"").encode("utf-8") for v in r])
        log_warn("   Engine Excel tidak tersedia/terkunci. Ringkasan disimpan dalam format CSV (bisa dibuka langsung di Excel): %s" % csv_path)
        return csv_path
    except Exception as e_csv:
        log_warn("   GAGAL total mengekspor ringkasan: %s" % str(e_csv))
        return None


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
        self.label = "3. QC Hasil Survey (Jumlah Sampel & RSD)"
        self.description = (
            "Membaca titik survei dari berbagai format Excel/CSV, mengisi otomatis parameter, "
            "menyamakan proyeksi, melakukan spatial join absolut ke poligon zona, "
            "serta menghitung RSD, status QC per zona, dan rekap surveyor tanpa distorsi ID."
        )
        self.canRunInBackground = True

    def getParameterInfo(self):
        p0 = arcpy.Parameter(
            displayName="File Input Survei (Excel .xlsx/.xls/.xlsm, CSV, TXT, atau Tabel)",
            name="input_file",
            datatype=["DEFile", "GPTableView"],
            parameterType="Required",
            direction="Input"
        )

        p1 = arcpy.Parameter(
            displayName="Nama Sheet Excel (otomatis terisi jika input file Excel)",
            name="sheet_name",
            datatype="GPString",
            parameterType="Optional",
            direction="Input"
        )
        p1.filter.type = "ValueList"
        p1.filter.list = []

        p2 = arcpy.Parameter(
            displayName="Kolom X (Longitude / Easting)",
            name="x_field",
            datatype="GPString",
            parameterType="Required",
            direction="Input"
        )
        p2.filter.type = "ValueList"
        p2.filter.list = []

        p3 = arcpy.Parameter(
            displayName="Kolom Y (Latitude / Northing)",
            name="y_field",
            datatype="GPString",
            parameterType="Required",
            direction="Input"
        )
        p3.filter.type = "ValueList"
        p3.filter.list = []

        p4 = arcpy.Parameter(
            displayName="Kolom Nilai (berisi nilai tanah per meter)",
            name="nilai_field",
            datatype="GPString",
            parameterType="Required",
            direction="Input"
        )
        p4.filter.type = "ValueList"
        p4.filter.list = []

        p5 = arcpy.Parameter(
            displayName="Surveyor/Penilai Tanah",
            name="surveyor_field",
            datatype="GPString",
            parameterType="Required",
            direction="Input"
        )
        p5.filter.type = "ValueList"
        p5.filter.list = []

        p6 = arcpy.Parameter(
            displayName="Shapefile Zona (polygon)",
            name="shp_zona",
            datatype="GPFeatureLayer",
            parameterType="Required",
            direction="Input"
        )
        p6.filter.list = ["Polygon"]

        p7 = arcpy.Parameter(
            displayName="NOZN (Kolom ID Zona)",
            name="id_zona_field",
            datatype="Field",
            parameterType="Required",
            direction="Input"
        )
        p7.parameterDependencies = [p6.name]

        p8 = arcpy.Parameter(
            displayName="Output Zona Hasil QC",
            name="output_fc",
            datatype="DEFeatureClass",
            parameterType="Required",
            direction="Output"
        )

        p9 = arcpy.Parameter(
            displayName="Output Titik Sampel",
            name="output_titik",
            datatype="DEFeatureClass",
            parameterType="Optional",
            direction="Output"
        )

        p10 = arcpy.Parameter(
            displayName="Skala Project",
            name="skala",
            datatype="GPString",
            parameterType="Required",
            direction="Input"
        )
        p10.filter.type = "ValueList"
        p10.filter.list = ["2500", "10000", "25000"]
        p10.value = "10000"

        p11 = arcpy.Parameter(
            displayName="Jumlah titik minimum per zona",
            name="min_titik",
            datatype="GPLong",
            parameterType="Required",
            direction="Input",
            category="Lanjutan"
        )
        p11.value = 3

        p12 = arcpy.Parameter(
            displayName="Sistem koordinat Excel/CSV (kosong = WGS 1984)",
            name="csv_sr",
            datatype="GPSpatialReference",
            parameterType="Optional",
            direction="Input",
            category="Lanjutan"
        )
        
        p13 = arcpy.Parameter(
            displayName="Output Excel Summary (Opsional)",
            name="out_excel",
            datatype="DEFile",
            parameterType="Optional",
            direction="Output"
        )

        p14 = arcpy.Parameter(
            displayName="Kolom Nomor/ID Sampel dari Excel (Opsional)",
            name="id_titik_field",
            datatype="GPString",
            parameterType="Optional",
            direction="Input",
            category="Lanjutan"
        )
        p14.filter.type = "ValueList"
        p14.filter.list = []

        return [p0, p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, p13, p14]

    def isLicensed(self):
        return True

    def updateParameters(self, parameters):
        in_path = parameters[0].valueAsText
        if in_path:
            file_path, sub_sheet = urai_path_excel(in_path)
            ext = os.path.splitext(file_path)[1].lower()

            if ext in (".xls", ".xlsx", ".xlsm", ".xlsb"):
                parameters[1].enabled = True
                sheets = daftar_sheet_excel(file_path)
                if sheets:
                    parameters[1].filter.list = sheets
                    if sub_sheet and sub_sheet in sheets and not parameters[1].altered:
                        parameters[1].value = sub_sheet
                    elif not parameters[1].valueAsText or parameters[1].valueAsText not in sheets:
                        parameters[1].value = sheets[0]
            else:
                parameters[1].enabled = False
                parameters[1].filter.list = []
                parameters[1].value = None

            chosen_sheet = parameters[1].valueAsText
            kolom_list = baca_daftar_kolom_input(in_path, chosen_sheet)

            if kolom_list:
                for idx in (2, 3, 4, 5, 14):
                    parameters[idx].filter.list = kolom_list
                    if parameters[idx].valueAsText and parameters[idx].valueAsText not in kolom_list:
                        parameters[idx].value = None

                if not parameters[2].valueAsText:
                    auto_x = tebak_kolom_terbaik(kolom_list, exact_candidates=["X", "KOORDINAT_X", "LONGITUDE", "EASTING"], keyword_candidates=["longitude", "easting"])
                    if auto_x: parameters[2].value = auto_x

                if not parameters[3].valueAsText:
                    auto_y = tebak_kolom_terbaik(kolom_list, exact_candidates=["Y", "KOORDINAT_Y", "LATITUDE", "NORTHING"], keyword_candidates=["latitude", "northing"])
                    if auto_y: parameters[3].value = auto_y

                if not parameters[4].valueAsText:
                    auto_val = tebak_kolom_terbaik(kolom_list, exact_candidates=["Nilai Tanah/m2", "Nilai_Tanah_m2", "NIR"], keyword_candidates=["nilaitanah", "harga"])
                    if auto_val: parameters[4].value = auto_val

                if not parameters[5].valueAsText:
                    auto_surv = tebak_kolom_terbaik(kolom_list, exact_candidates=["Nama Surveyor", "Surveyor", "Penilai"], keyword_candidates=["surveyor", "penilai"])
                    if auto_surv: parameters[5].value = auto_surv

                if not parameters[14].valueAsText:
                    auto_id_sampel = tebak_kolom_terbaik(
                        kolom_list, 
                        exact_candidates=["ID_TITIK", "NO_SAMPEL", "ID_SAMPEL", "NO", "NOMOR", "FID_TITIK", "ID"], 
                        keyword_candidates=["idsampel", "idtitik", "nomor"]
                    )
                    if auto_id_sampel:
                        parameters[14].value = auto_id_sampel

        shp_zona = parameters[6].valueAsText
        if shp_zona:
            try:
                zona_fields = [f.name for f in arcpy.ListFields(shp_zona) if f.type not in ("OID", "Geometry")]
                if zona_fields:
                    curr_id_z = parameters[7].valueAsText
                    if not curr_id_z or curr_id_z not in zona_fields:
                        auto_nozn = tebak_kolom_terbaik(zona_fields, exact_candidates=["NOZN", "ID_ZONA", "ZONA"], keyword_candidates=["nozn", "kodezona"])
                        if auto_nozn: parameters[7].value = auto_nozn
            except Exception:
                pass

        if parameters[13].altered and parameters[13].valueAsText:
            raw_ex = parameters[13].valueAsText
            ref_fc = parameters[8].valueAsText
            fixed_ex = normalisasi_path_output_excel(raw_ex, ref_fc)
            if fixed_ex and fixed_ex != raw_ex:
                parameters[13].value = fixed_ex
        return

    def updateMessages(self, parameters):
        if parameters[11].value is not None and parameters[11].value < 2:
            parameters[11].setErrorMessage("Minimum titik harus >= 2 agar RSD bisa dihitung")
        if parameters[2].value and parameters[3].value and parameters[2].valueAsText == parameters[3].valueAsText:
            parameters[3].setErrorMessage("Kolom Y tidak boleh sama dengan kolom X")
        return

    def execute(self, parameters, messages):
        in_file = parameters[0].valueAsText
        sheet_name = parameters[1].valueAsText
        x_field = parameters[2].valueAsText
        y_field = parameters[3].valueAsText
        nilai_field = parameters[4].valueAsText
        surveyor_field = parameters[5].valueAsText
        shp_zona = parameters[6].valueAsText
        id_zona_field = parameters[7].valueAsText
        output_fc = parameters[8].valueAsText
        output_titik = parameters[9].valueAsText
        skala = unicode(parameters[10].valueAsText).strip()
        min_titik = int(parameters[11].value)
        csv_sr = parameters[12].value or arcpy.SpatialReference(4326)
        out_excel = parameters[13].valueAsText
        id_titik_field = parameters[14].valueAsText 

        max_rsd = {"2500": 15.0, "10000": 25.0, "25000": 30.0}.get(skala, 25.0)

        arcpy.env.overwriteOutput = True
        arcpy.env.qualifiedFieldNames = False
        scratch = arcpy.env.scratchGDB
        if not scratch or not arcpy.Exists(scratch):
            scratch = arcpy.env.scratchFolder or os.path.dirname(output_fc)

        temp_table = os.path.join(scratch, "qc_temp_table")
        temp_layer = "csv_virtual_layer"
        temp_points = os.path.join(scratch, "qc_temp_points")
        temp_proj = os.path.join(scratch, "qc_temp_proj")
        temp_join = os.path.join(scratch, "qc_temp_join")

        def msg(t):
            messages.addMessage(t)

        def gagal(t):
            arcpy.AddError(t)
            raise ValueError(t)

        try:
            msg("0. Memeriksa sistem koordinat...")
            zona_sr = arcpy.Describe(shp_zona).spatialReference
            if sr_tidak_dikenal(zona_sr):
                gagal("Shapefile zona tidak punya sistem koordinat (Unknown).")
            if sr_tidak_dikenal(csv_sr):
                gagal("Sistem koordinat Excel/CSV tidak valid/Unknown.")
            
            msg("1. Menyiapkan poligon zona dan mengamankan Spatial ID...")
            arcpy.CopyFeatures_management(shp_zona, output_fc)
            
            kolom_zona_spasial = "ID_ZN_SPAT"
            arcpy.AddField_management(output_fc, kolom_zona_spasial, "TEXT", field_length=150)
            arcpy.CalculateField_management(output_fc, kolom_zona_spasial, "!{}!".format(id_zona_field), "PYTHON_9.3")

            msg("2. Membaca file input Excel dan membangun geometri titik...")
            if arcpy.Exists(temp_table):
                arcpy.Delete_management(temp_table)

            siapkan_tabel_gdb_terkunci(
                in_file, sheet_name, scratch, "qc_temp_table",
                x_field, y_field, nilai_field, surveyor_field, messages
            )
            cek_rentang_xy(temp_table, "QC_X_NUM", "QC_Y_NUM", csv_sr, messages)

            if arcpy.Exists(temp_layer):
                arcpy.Delete_management(temp_layer)
            arcpy.MakeXYEventLayer_management(temp_table, "QC_X_NUM", "QC_Y_NUM", temp_layer, csv_sr)

            if arcpy.Exists(temp_points):
                arcpy.Delete_management(temp_points)
            arcpy.CopyFeatures_management(temp_layer, temp_points)

            n_pt = int(arcpy.GetCount_management(temp_points).getOutput(0))
            if n_pt == 0:
                gagal("Tidak ada titik terbentuk. Cek kolom X/Y pada file input.")

            # =========================================================
            # MENGUNCI NOMOR SAMPEL AGAR ANTI-GESER
            # =========================================================
            msg("   Mengamankan Nomor Sampel asli dari file Excel...")
            kolom_sampel_aman = "ID_SMPL" 
            arcpy.AddField_management(temp_points, kolom_sampel_aman, "TEXT", field_length=50)
            
            field_points = [f.name for f in arcpy.ListFields(temp_points)]
            col_id_titik_aktual = cari_nama_kolom_ketat(field_points, id_titik_field, wajib=False) if id_titik_field else None

            if col_id_titik_aktual:
                arcpy.CalculateField_management(temp_points, kolom_sampel_aman, "!{}!".format(col_id_titik_aktual), "PYTHON_9.3")
            else:
                code_block = "rec = 0\ndef get_id():\n    global rec\n    rec += 1\n    return str(rec)"
                arcpy.CalculateField_management(temp_points, kolom_sampel_aman, "get_id()", "PYTHON_9.3", code_block)
            # =========================================================

            titik_siap = temp_points
            if not sr_sama(csv_sr, zona_sr):
                msg("   Memproyeksikan titik ke CRS zona...")
                transf = pilih_transformasi(csv_sr, zona_sr, temp_points)
                if arcpy.Exists(temp_proj):
                    arcpy.Delete_management(temp_proj)
                if transf:
                    arcpy.Project_management(temp_points, temp_proj, zona_sr, transf)
                else:
                    arcpy.Project_management(temp_points, temp_proj, zona_sr)
                titik_siap = temp_proj

            cek_overlap(titik_siap, output_fc, messages)

            msg("3. Spatial join titik ke zona (berdasarkan letak spasial murni)...")
            if arcpy.Exists(temp_join):
                arcpy.Delete_management(temp_join)

            arcpy.SpatialJoin_analysis(
                target_features=titik_siap,
                join_features=output_fc,
                out_feature_class=temp_join,
                join_operation="JOIN_ONE_TO_MANY",
                join_type="KEEP_ALL",
                match_option="INTERSECT"
            )

            if output_titik:
                msg("   Menyimpan titik hasil join...")
                simpan_titik_hasil_join(temp_join, output_titik)
                arcpy.SetParameterAsText(9, output_titik)

            msg("4. Menghitung RSD dan merekap surveyor per zona...")
            actual = [f.name for f in arcpy.ListFields(temp_join)]
            
            col_id_actual = cari_nama_kolom_ketat(actual, kolom_zona_spasial)
            col_nilai_actual = cari_nama_kolom_ketat(actual, "QC_VAL_NUM")
            col_surv_actual = cari_nama_kolom_ketat(actual, "QC_SURV_TXT")
            col_sampel_actual = cari_nama_kolom_ketat(actual, "ID_SMPL")

            zona_data = {}
            n_nilai_invalid = 0

            cursor_fields = [col_id_actual, col_nilai_actual, col_surv_actual, col_sampel_actual]
            
            with arcpy.da.SearchCursor(temp_join, cursor_fields) as cur:
                for row_j in cur:
                    id_z = row_j[0]       
                    val = row_j[1]
                    surv = row_j[2]
                    id_sampel = row_j[3]  

                    norm_id = normalisasi_id_zona(id_z)
                    if norm_id is None:
                        continue

                    d = zona_data.setdefault(norm_id, {"nilai": [], "surveyor": set(), "sampel": []})

                    v = bersihkan_angka(val)
                    if v is not None:
                        d["nilai"].append(v)
                        d["sampel"].append(unicode(id_sampel) if id_sampel is not None else "?")
                    else:
                        n_nilai_invalid += 1

                    if surv is not None and unicode(surv).strip() not in ("", "-", "None", "NULL"):
                        d["surveyor"].add(unicode(surv).strip())

            if n_nilai_invalid:
                messages.addWarningMessage(
                    "   %d titik di dalam zona memiliki nilai kosong/bukan angka pada kolom '%s'."
                    % (n_nilai_invalid, nilai_field)
                )

            msg("   Total zona yang memiliki titik survei: %d zona" % len(zona_data))

            msg("5. Mengkalkulasi Atribut QC Final...")
            out_fields = [f.name for f in arcpy.ListFields(output_fc)]
            col_id_out = cari_nama_kolom_ketat(out_fields, id_zona_field)

            qc_field_specs = [
                ("Jml_Titik", "LONG", None),
                ("Rata_Nilai", "DOUBLE", None),
                ("RSD_Persen", "DOUBLE", None),
                ("Status_QC", "TEXT", 100),
                ("Surveyor", "TEXT", 254)
            ]
            for fname, ftype, flen in qc_field_specs:
                existing_match = cari_nama_kolom_ketat([f.name for f in arcpy.ListFields(output_fc)], fname, wajib=False)
                if existing_match and existing_match.upper() != col_id_out.upper():
                    try:
                        arcpy.DeleteField_management(output_fc, existing_match)
                    except Exception:
                        pass
                existing_now = [f.name.upper() for f in arcpy.ListFields(output_fc)]
                if fname.upper() not in existing_now and fname[:10].upper() not in existing_now:
                    if flen:
                        arcpy.AddField_management(output_fc, fname, ftype, field_length=flen)
                    else:
                        arcpy.AddField_management(output_fc, fname, ftype)

            if kolom_zona_spasial in [f.name for f in arcpy.ListFields(output_fc)]:
                arcpy.DeleteField_management(output_fc, kolom_zona_spasial)

            out_fields_updated = [f.name for f in arcpy.ListFields(output_fc)]
            f_jml = cari_nama_kolom_ketat(out_fields_updated, "Jml_Titik")
            f_rata = cari_nama_kolom_ketat(out_fields_updated, "Rata_Nilai")
            f_rsd = cari_nama_kolom_ketat(out_fields_updated, "RSD_Persen")
            f_stat = cari_nama_kolom_ketat(out_fields_updated, "Status_QC")
            f_surv = cari_nama_kolom_ketat(out_fields_updated, "Surveyor")

            n_valid = n_revisi = n_kurang = n_kosong = 0
            upd_fields = [col_id_out, f_jml, f_rata, f_rsd, f_stat, f_surv]
            excel_rows_summary = []

            with arcpy.da.UpdateCursor(output_fc, upd_fields) as cur:
                for row in cur:
                    norm_z = normalisasi_id_zona(row[0])
                    d = zona_data.get(norm_z) if norm_z is not None else None

                    if not d or len(d["nilai"]) == 0:
                        cur.updateRow([row[0], 0, 0.0, 0.0, "Belum Disurvey", "-"])
                        n_kosong += 1
                        continue

                    vals = d["nilai"]
                    n = len(vals)
                    rata2 = float(sum(vals)) / float(n)
                    rsd = 0.0
                    
                    if n > 1 and rata2 != 0.0:
                        var = sum((float(x) - rata2) ** 2 for x in vals) / float(n - 1)
                        std_dev = math.sqrt(var)
                        rsd = (std_dev / abs(rata2)) * 100.0

                    rata2_round = round(rata2, 2)
                    rsd_round = round(rsd, 2)

                    if n < min_titik:
                        status = "Kurang Lengkap (Tunggu min %d titik)" % min_titik
                        n_kurang += 1
                    elif rsd_round == 0.0:
                        status = "Perlu Revisi (RSD 0% tidak diizinkan)"
                        n_revisi += 1
                    elif rsd_round > max_rsd:
                        status = "Perlu Revisi (RSD %.2f%% > batas %d%%)" % (rsd_round, int(max_rsd))
                        n_revisi += 1
                    else:
                        status = "Lengkap & Valid"
                        n_valid += 1

                    surv_str = ", ".join(sorted(list(d["surveyor"]))) or "-"
                    surv_str = surv_str[:254]
                    status_str = status[:100]

                    cur.updateRow([row[0], int(n), rata2_round, rsd_round, status_str, surv_str])

                    daftar_sampel = ", ".join(d["sampel"]) if d.get("sampel") else "-"
                    no_zona_str = unicode(row[0]) if row[0] is not None else u"-"
                    
                    excel_rows_summary.append([
                        no_zona_str,
                        int(n),
                        daftar_sampel,
                        float(rata2_round),
                        float(rsd_round),
                        unicode(status_str),
                        unicode(surv_str)
                    ])

            msg("Rekap zona: %d valid | %d perlu revisi | %d kurang lengkap | %d belum disurvey"
                % (n_valid, n_revisi, n_kurang, n_kosong))
            arcpy.SetParameterAsText(8, output_fc)

            # --- MULAI EKSPOR EXCEL ---
            if out_excel:
                msg("6. Mengekspor Ringkasan QC ke Excel (mengabaikan zona Belum Disurvey)...")
                excel_rows_summary.sort(key=lambda r: natural_key(unicode(r[0])))

                excel_headers = ["No_Zona", "Jml_Sampel", "Sampel_QC", "Rata_Nilai", "RSD_Persen", "Status_QC", "Surveyor"]
                excel_specs = [
                    ("No_Zona", "TEXT", 150),
                    ("Jml_Sampel", "LONG", None),
                    ("Sampel_QC", "TEXT", 254),
                    ("Rata_Nilai", "DOUBLE", None),
                    ("RSD_Persen", "DOUBLE", None),
                    ("Status_QC", "TEXT", 150),
                    ("Surveyor", "TEXT", 250)
                ]

                saved_excel = ekspor_data_ke_excel(
                    headers=excel_headers,
                    rows=excel_rows_summary,
                    field_specs=excel_specs,
                    out_excel_raw=out_excel,
                    fallback_ref_path=output_fc,
                    messages=messages
                )
                if saved_excel:
                    arcpy.SetParameterAsText(13, saved_excel)
            # --- AKHIR EKSPOR EXCEL ---

            msg("Proses selesai.")

        finally:
            for f in (temp_layer, temp_table, temp_points, temp_proj, temp_join):
                try:
                    if arcpy.Exists(f):
                        arcpy.Delete_management(f)
                except Exception:
                    pass
        return

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


# =====================================================================
# TOOL 5: EKSTRAKSI ZONA TERTINGGI & TERENDAH KE EXCEL (SumZNT)
# =====================================================================
class EkstrakZonaExcel(object):
    def __init__(self):
        self.label = "5. (Final Report) Ekstrak Nilai ZNT Max & Min ke Excel"
        self.description = ("Ekstraksi nilai ZNT tertinggi dan terendah per KECAMATAN beserta info Zona & Desa "
                            "ke Excel. Berjalan di ArcGIS Pro dan ArcGIS Desktop.")
        self.canRunInBackground = False

    def getParameterInfo(self):
        p_zona = arcpy.Parameter(name="in_zona", displayName="1. Layer ZONA",
                                 datatype="GPFeatureLayer", parameterType="Required", direction="Input")

        p_fld_zona = arcpy.Parameter(name="field_zona", displayName="   -> Field Nomor Zona (Contoh: NOZN)",
                                     datatype="Field", parameterType="Required", direction="Input")
        p_fld_zona.parameterDependencies = [p_zona.name]

        p_fld_harga = arcpy.Parameter(name="field_harga", displayName="   -> Field Harga/Nilai (Contoh: RPBULAT)",
                                      datatype="Field", parameterType="Required", direction="Input")
        p_fld_harga.parameterDependencies = [p_zona.name]

        p_metode = arcpy.Parameter(name="pilih_metode", displayName="2. Pilih Metode Analisis",
                                   datatype="GPString", parameterType="Required", direction="Input")
        p_metode.filter.type = "ValueList"
        p_metode.filter.list = [
            "Metode Centroid",
            "Metode Intersect"]
        p_metode.value = p_metode.filter.list[0]

        p_kec = arcpy.Parameter(name="in_kecamatan", displayName="3. Batas KECAMATAN (SHP / Geodatabase)",
                                datatype="GPFeatureLayer", parameterType="Required", direction="Input")

        p_fld_kec = arcpy.Parameter(name="field_kecamatan", displayName="   -> Field Nama Kecamatan",
                                    datatype="Field", parameterType="Required", direction="Input")
        p_fld_kec.parameterDependencies = [p_kec.name]

        p_enable_desa = arcpy.Parameter(name="enable_desa", displayName="Sertakan Analisis Level DESA",
                                        datatype="GPBoolean", parameterType="Optional", direction="Input")
        p_enable_desa.value = True

        p_desa = arcpy.Parameter(name="in_desa", displayName="4. Batas DESA (SHP / Geodatabase)",
                                 datatype="GPFeatureLayer", parameterType="Optional", direction="Input")

        p_fld_desa = arcpy.Parameter(name="field_desa", displayName="   -> Field Nama Desa",
                                     datatype="Field", parameterType="Optional", direction="Input")
        p_fld_desa.parameterDependencies = [p_desa.name]

        p_min_area = arcpy.Parameter(name="min_area",
                                     displayName="5. Abaikan fragmen irisan lebih kecil dari (m2) - khusus Metode Intersect",
                                     datatype="GPDouble", parameterType="Optional", direction="Input")
        p_min_area.value = 100

        p_out = arcpy.Parameter(name="out_excel", displayName="6. Output File Excel",
                                datatype="DEFile", parameterType="Required", direction="Output")

        return [p_zona, p_fld_zona, p_fld_harga, p_metode,
                p_kec, p_fld_kec, p_enable_desa, p_desa, p_fld_desa,
                p_min_area, p_out]

    def isLicensed(self):
        return True

    def updateParameters(self, parameters):
        enabled = bool(parameters[6].value)
        parameters[7].enabled = enabled
        parameters[8].enabled = enabled
        metode = parameters[3].valueAsText or ""
        parameters[9].enabled = ("Centroid" not in metode)

        # Auto-fill NOZN dan RPBULAT/Rata_Nilai pada Layer Zona
        if parameters[0].valueAsText:
            try:
                z_fields = [
                    f.name for f in arcpy.ListFields(parameters[0].valueAsText)
                    if f.type not in ("OID", "Geometry")
                ]
                if z_fields:
                    if not parameters[1].valueAsText or parameters[1].valueAsText not in z_fields:
                        auto_z = tebak_kolom_terbaik(z_fields, ["NOZN", "NO_ZONA", "ID_ZONA", "ZONA"])
                        if auto_z:
                            parameters[1].value = auto_z
                    if not parameters[2].valueAsText or parameters[2].valueAsText not in z_fields:
                        auto_h = tebak_kolom_terbaik(z_fields, ["RPBULAT", "NIR", "Rata_Nilai", "NILAI", "HARGA"])
                        if auto_h:
                            parameters[2].value = auto_h
            except Exception:
                pass

        if parameters[10].altered and parameters[10].valueAsText:
            raw_ex = parameters[10].valueAsText
            fixed_ex = normalisasi_path_output_excel(raw_ex)
            if fixed_ex and fixed_ex != raw_ex:
                parameters[10].value = fixed_ex
        return

    def updateMessages(self, parameters):
        if parameters[6].value:
            if not parameters[7].value:
                parameters[7].setErrorMessage(
                    "Pilih layer Batas DESA, atau matikan 'Sertakan Analisis Level DESA'.")
            elif not parameters[8].value:
                parameters[8].setErrorMessage("Pilih field nama desa.")
        if parameters[9].value is not None and parameters[9].value < 0:
            parameters[9].setErrorMessage("Nilai tidak boleh negatif.")
        return

    def execute(self, parameters, messages):
        arcpy.env.overwriteOutput = True
        scratch_db = arcpy.env.scratchGDB
        arcpy.env.workspace = scratch_db

        def tmp_path(name):
            return os.path.join(scratch_db, name)

        zona_pt = tmp_path("zona_titik_tmp")
        join_1 = tmp_path("join_kecamatan_tmp")
        join_2 = tmp_path("join_desa_tmp")
        z_slim = tmp_path("zona_slim_tmp")
        k_slim = tmp_path("kec_slim_tmp")
        d_slim = tmp_path("desa_slim_tmp")
        inter_k = tmp_path("intersect_kec_tmp")
        inter_kd = tmp_path("intersect_desa_tmp")
        temp_table = tmp_path("tabel_export_tmp")
        temps = [zona_pt, join_1, join_2, z_slim, k_slim, d_slim,
                 inter_k, inter_kd, temp_table]

        try:
            in_zona = parameters[0].valueAsText
            fld_zona_in = parameters[1].valueAsText
            fld_harga_in = parameters[2].valueAsText
            metode = parameters[3].valueAsText
            in_kec = parameters[4].valueAsText
            fld_kec_in = parameters[5].valueAsText
            enable_desa = bool(parameters[6].value)
            in_desa = parameters[7].valueAsText if enable_desa else None
            fld_desa_in = parameters[8].valueAsText if enable_desa else None
            min_area = float(parameters[9].value or 0)
            out_excel = parameters[10].valueAsText

            if enable_desa and (not in_desa or not fld_desa_in):
                raise Exception("Analisis desa aktif tetapi layer/field desa belum dipilih.")

            use_closest = "Centroid" in metode

            info = arcpy.GetInstallInfo()
            arcpy.AddMessage("Aplikasi: {0} {1} | Python {2}".format(
                info.get("ProductName", "ArcGIS"), info.get("Version", ""),
                sys.version.split()[0]))

            for tmp in temps:
                if arcpy.Exists(tmp):
                    arcpy.Delete_management(tmp)

            arcpy.AddMessage("1/4: Memeriksa layer zona...")
            n_zona = int(arcpy.GetCount_management(in_zona).getOutput(0))
            arcpy.AddMessage("     -> Fitur zona yang diproses: {0}".format(n_zona))
            if n_zona == 0:
                raise Exception("Layer zona kosong. Cek seleksi aktif atau definition query.")

            sr_zona = arcpy.Describe(in_zona).spatialReference
            if sr_zona.name == "Unknown":
                arcpy.AddWarning("Layer zona tidak punya sistem koordinat (Unknown).")
            arcpy.env.outputCoordinateSystem = sr_zona
            arcpy.AddMessage("     -> CRS eksekusi: {0}".format(sr_zona.name))

            stats = _sum_new_stats()
            records = []
            desa_lookup = None

            if use_closest:
                arcpy.AddMessage("2/4: Memproses dengan Titik Centroid (CLOSEST)...")
                arcpy.FeatureToPoint_management(in_zona, zona_pt, "INSIDE")
                base_fc = zona_pt

                fld_zona = resolve_field(base_fc, fld_zona_in)
                fld_harga = resolve_field(base_fc, fld_harga_in)
                kec_src = resolve_field(in_kec, fld_kec_in)

                fm1 = _sum_make_field_mappings(base_fc, [fld_zona, fld_harga], in_kec, kec_src, "KEC_OUT")
                arcpy.SpatialJoin_analysis(base_fc, in_kec, join_1,
                                           "JOIN_ONE_TO_ONE", "KEEP_ALL", fm1, "CLOSEST")
                result_fc = join_1

                if enable_desa:
                    desa_src = resolve_field(in_desa, fld_desa_in)
                    keep2 = [resolve_field(join_1, fld_zona),
                             resolve_field(join_1, fld_harga),
                             resolve_field(join_1, "KEC_OUT")]
                    fm2 = _sum_make_field_mappings(join_1, keep2, in_desa, desa_src, "DESA_OUT")
                    arcpy.SpatialJoin_analysis(join_1, in_desa, join_2,
                                               "JOIN_ONE_TO_ONE", "KEEP_ALL", fm2, "CLOSEST")
                    result_fc = join_2

                arcpy.AddMessage("     -> Baris hasil proses spasial: {0}".format(
                    int(arcpy.GetCount_management(result_fc).getOutput(0))))

                arcpy.AddMessage("3/4: Membaca & mengagregasi per KECAMATAN...")
                f_zona = resolve_field(result_fc, fld_zona_in)
                f_harga = resolve_field(result_fc, fld_harga_in)
                f_kec = resolve_field(result_fc, "KEC_OUT")
                read_fields = [f_zona, f_harga, f_kec]
                if enable_desa:
                    read_fields.append(resolve_field(result_fc, "DESA_OUT"))

                with arcpy.da.SearchCursor(result_fc, read_fields) as cursor:
                    for row in cursor:
                        stats["total"] += 1
                        ok = _sum_clean_row(row[0], row[1], stats)
                        if ok is None:
                            continue
                        zona, harga = ok
                        kec = to_text(row[2]).strip() or TIDAK_TERIDENTIFIKASI
                        if enable_desa:
                            desa = to_text(row[3]).strip() or TIDAK_TERIDENTIFIKASI
                        else:
                            desa = "-"
                        stats["valid"] += 1
                        records.append((kec, zona, harga, desa))

            else:
                arcpy.AddMessage("2/4: Intersect klasik (poligon zona dipotong pada batas wilayah)...")
                geodesic = (sr_zona.type != "Projected")
                if geodesic:
                    arcpy.AddWarning("CRS zona geografis (derajat). Luas dihitung secara geodesik; disarankan memakai CRS terproyeksi.")

                col_z = resolve_field(in_zona, fld_zona_in)
                col_h = resolve_field(in_zona, fld_harga_in)
                col_k = resolve_field(in_kec, fld_kec_in)

                arcpy.AddMessage("     -> Menyiapkan salinan ringkas...")
                _sum_slim_copy(in_zona, z_slim, [(col_z, "ZONA_OUT"), (col_h, "HARGA_OUT")], sr_zona)
                _sum_slim_copy(in_kec, k_slim, [(col_k, "KEC_OUT")], sr_zona)

                arcpy.Intersect_analysis([z_slim, k_slim], inter_k, "NO_FID")
                n_frag = int(arcpy.GetCount_management(inter_k).getOutput(0))
                arcpy.AddMessage("     -> Fragmen irisan zona x kecamatan: {0}".format(n_frag))

                z_total = 0.0
                with arcpy.da.SearchCursor(z_slim, ["SHAPE@"]) as cursor:
                    for (shp,) in cursor:
                        if shp is not None:
                            z_total += _sum_area_m2(shp, geodesic)

                arcpy.AddMessage("3/4: Membaca & mengagregasi per KECAMATAN...")
                fz = resolve_field(inter_k, "ZONA_OUT")
                fh = resolve_field(inter_k, "HARGA_OUT")
                fk = resolve_field(inter_k, "KEC_OUT")
                frag_total = 0.0
                with arcpy.da.SearchCursor(inter_k, ["SHAPE@", fz, fh, fk]) as cursor:
                    for shp, z_raw, h_raw, k_raw in cursor:
                        stats["total"] += 1
                        a = _sum_area_m2(shp, geodesic) if shp is not None else 0.0
                        frag_total += a
                        if a < min_area:
                            stats["fragmen_kecil"] += 1
                            continue
                        ok = _sum_clean_row(z_raw, h_raw, stats)
                        if ok is None:
                            continue
                        zona, harga = ok
                        kec = to_text(k_raw).strip() or TIDAK_TERIDENTIFIKASI
                        stats["valid"] += 1
                        records.append((kec, zona, harga, None))

                if z_total > 0:
                    cakupan = 100.0 * frag_total / z_total
                    arcpy.AddMessage("     -> Cakupan: {0:.1f}% luas zona berada di dalam batas kecamatan".format(cakupan))

                if enable_desa:
                    arcpy.AddMessage("     -> Intersect zona x kecamatan x desa...")
                    desa_src = resolve_field(in_desa, fld_desa_in)
                    _sum_slim_copy(in_desa, d_slim, [(desa_src, "DESA_OUT")], sr_zona)
                    arcpy.Intersect_analysis([z_slim, k_slim, d_slim], inter_kd, "NO_FID")

                    gz = resolve_field(inter_kd, "ZONA_OUT")
                    gh = resolve_field(inter_kd, "HARGA_OUT")
                    gk = resolve_field(inter_kd, "KEC_OUT")
                    gd = resolve_field(inter_kd, "DESA_OUT")
                    desa_lookup = {}
                    with arcpy.da.SearchCursor(inter_kd, ["SHAPE@", gz, gh, gk, gd]) as cursor:
                        for shp, z_raw, h_raw, k_raw, d_raw in cursor:
                            if shp is None or _sum_area_m2(shp, geodesic) < min_area:
                                continue
                            zona = to_text(z_raw).strip()
                            harga = parse_number(h_raw)
                            if not zona or harga is None:
                                continue
                            kec = to_text(k_raw).strip() or TIDAK_TERIDENTIFIKASI
                            desa = to_text(d_raw).strip() or TIDAK_TERIDENTIFIKASI
                            desa_lookup.setdefault((kec, zona, harga), set()).add(desa)

            arcpy.AddMessage("     -> Ringkasan baca: total={total}, valid={valid}, "
                             "zona kosong={zona_kosong}, harga kosong={harga_kosong}, "
                             "harga bukan angka={harga_bukan_angka}, "
                             "fragmen < {0} m2 dibuang={fragmen_kecil}".format(min_area, **stats))

            if stats["valid"] == 0 or not records:
                raise Exception("Tidak ada baris valid. Periksa field harga/zona, CRS layer, dsb.")

            rows_to_insert = _sum_aggregate_by_kecamatan(records, desa_lookup, enable_desa)

            arcpy.AddMessage("4/4: Menyusun tabel Excel fokus per KECAMATAN...")
            sum_headers = [
                "KECAMATAN", "ZNT_TERENDAH", "ZONA_TERENDAH", "DESA_TERENDAH",
                "ZNT_TERTINGGI", "ZONA_TERTINGGI", "DESA_TERTINGGI"
            ]
            sum_specs = [
                ("KECAMATAN", "TEXT", 150),
                ("ZNT_TERENDAH", "DOUBLE", None),
                ("ZONA_TERENDAH", "TEXT", 254),
                ("DESA_TERENDAH", "TEXT", 254),
                ("ZNT_TERTINGGI", "DOUBLE", None),
                ("ZONA_TERTINGGI", "TEXT", 254),
                ("DESA_TERTINGGI", "TEXT", 254)
            ]

            saved_excel = ekspor_data_ke_excel(
                headers=sum_headers,
                rows=rows_to_insert,
                field_specs=sum_specs,
                out_excel_raw=out_excel,
                messages=messages
            )
            if saved_excel:
                arcpy.SetParameterAsText(10, saved_excel)

            arcpy.AddMessage("=== PROSES SELESAI ===")
            arcpy.AddMessage("{0} baris Kecamatan ditulis ke: {1}".format(len(rows_to_insert), saved_excel or out_excel))

        except Exception as e:
            tb = sys.exc_info()[2]
            line_num = traceback.tb_lineno(tb)
            arcpy.AddError("TERJADI KESALAHAN (Baris {0}): {1}".format(line_num, e))
        finally:
            for tmp in temps:
                try:
                    if arcpy.Exists(tmp):
                        arcpy.Delete_management(tmp)
                except Exception:
                    pass
        return


# =============================================================================
# FUNGSI PEMBACA MULTI-FORMAT EXCEL (.XLSX, .XLS, .XLSM) & PENGUNCI KOLOM
# (DIBAWA OLEH QC SURVEI ZONA)
# =============================================================================

def urai_path_excel(in_path):
    if not in_path:
        return "", None
    clean_p = in_path.strip("'\"")
    low_p = clean_p.lower()
    for ext in (".xlsx", ".xlsm", ".xlsb", ".xls"):
        idx = low_p.find(ext)
        if idx != -1:
            akhir_ext = idx + len(ext)
            file_part = clean_p[:akhir_ext]
            rest = clean_p[akhir_ext:].strip("\\/")
            if rest:
                sheet_part = rest.strip("'\"").rstrip("$")
                return file_part, sheet_part
            return file_part, None
    return clean_p, None

def daftar_sheet_excel(excel_path):
    if not os.path.isfile(excel_path):
        return []
    try:
        import xlrd
        wb = xlrd.open_workbook(excel_path, on_demand=True)
        names = [unicode(s) for s in wb.sheet_names()]
        wb.release_resources()
        if names: return names
    except Exception:
        pass
    try:
        import openpyxl
        wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
        names = [unicode(s) for s in wb.sheetnames]
        wb.close()
        if names: return names
    except Exception:
        pass
    try:
        old_ws = arcpy.env.workspace
        arcpy.env.workspace = excel_path
        tbls = arcpy.ListTables() or []
        arcpy.env.workspace = old_ws
        names = []
        for t in tbls:
            s = t.strip("'\"").rstrip("$")
            if s and not s.endswith("_"):
                names.append(s)
        return names
    except Exception:
        return []

def buat_header_unik(raw_headers):
    headers = []
    seen = {}
    for i, h in enumerate(raw_headers):
        s = unicode(h).strip() if h is not None else ""
        if not s: s = "COL_%d" % (i + 1)
        if s in seen:
            seen[s] += 1
            s = "%s_%d" % (s, seen[s])
        else:
            seen[s] = 1
        headers.append(s)
    return headers

def baca_data_excel_python(excel_path, sheet_name=None, hanya_header=False):
    try:
        import xlrd
        wb = xlrd.open_workbook(excel_path)
        if sheet_name and sheet_name in wb.sheet_names():
            sh = wb.sheet_by_name(sheet_name)
        else:
            sh = wb.sheet_by_index(0)

        if sh.nrows == 0: return [], []
        header_row_idx = 0
        for r_idx in range(min(10, sh.nrows)):
            row_vals = [unicode(sh.cell_value(r_idx, c)).strip() for c in range(sh.ncols)]
            if sum(1 for v in row_vals if v != "") >= 2:
                header_row_idx = r_idx
                break

        raw_headers = [sh.cell_value(header_row_idx, c) for c in range(sh.ncols)]
        headers = buat_header_unik(raw_headers)

        if hanya_header: return headers, []

        rows = []
        for r in range(header_row_idx + 1, sh.nrows):
            row_data = []
            is_empty = True
            for c in range(sh.ncols):
                cell = sh.cell(r, c)
                val = cell.value
                if cell.ctype == xlrd.XL_CELL_EMPTY: val = None
                elif cell.ctype == xlrd.XL_CELL_NUMBER:
                    if val == int(val): val = int(val)
                elif cell.ctype == xlrd.XL_CELL_DATE:
                    try:
                        dt = xlrd.xldate_as_tuple(val, wb.datemode)
                        val = "%04d-%02d-%02d" % (dt[0], dt[1], dt[2])
                    except Exception:
                        val = unicode(val)
                else:
                    val = unicode(val).strip()
                    if val == "": val = None
                if val is not None: is_empty = False
                row_data.append(val)
            if not is_empty:
                rows.append(row_data)
        return headers, rows
    except Exception:
        pass

    try:
        import openpyxl
        wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
        if sheet_name and sheet_name in wb.sheetnames:
            sh = wb[sheet_name]
        else:
            sh = wb[wb.sheetnames[0]]

        all_rows = list(sh.iter_rows(values_only=True))
        wb.close()
        if not all_rows: return [], []

        header_row_idx = 0
        for r_idx in range(min(10, len(all_rows))):
            row_vals = [unicode(v).strip() if v is not None else "" for v in all_rows[r_idx]]
            if sum(1 for v in row_vals if v != "") >= 2:
                header_row_idx = r_idx
                break

        headers = buat_header_unik(all_rows[header_row_idx])
        if hanya_header: return headers, []

        rows = []
        for r in all_rows[header_row_idx + 1:]:
            row_data = []
            is_empty = True
            for c in range(len(headers)):
                val = r[c] if c < len(r) else None
                if isinstance(val, basestring):
                    val = val.strip()
                    if val == "": val = None
                if val is not None: is_empty = False
                row_data.append(val)
            if not is_empty:
                rows.append(row_data)
        return headers, rows
    except Exception:
        return [], []

def baca_data_csv_python(csv_path, hanya_header=False):
    if not os.path.isfile(csv_path): return [], []
    try:
        with open(csv_path, "rb") as f:
            raw = f.read()
        for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
            try:
                text = raw.decode(enc)
                break
            except Exception: continue
        else:
            text = raw.decode("latin-1", errors="ignore")

        lines = [l for l in text.splitlines() if l.strip()]
        if not lines: return [], []

        sample = "\n".join(lines[:10])
        delim_candidates = [",", ";", "\t", "|"]
        delim = max(delim_candidates, key=lambda d: lines[0].count(d))
        if lines[0].count(delim) == 0:
            try:
                sniff = csv.Sniffer().sniff(sample)
                delim = sniff.delimiter
            except Exception:
                delim = ","

        if sys.version_info[0] < 3:
            encoded_lines = [l.encode("utf-8") for l in lines]
            reader = csv.reader(encoded_lines, delimiter=str(delim))
            all_rows = [[cell.decode("utf-8").strip() for cell in row] for row in reader]
        else:
            reader = csv.reader(lines, delimiter=delim)
            all_rows = [[cell.strip() for cell in row] for row in reader]

        if not all_rows: return [], []

        headers = buat_header_unik(all_rows[0])
        if hanya_header: return headers, []

        rows = []
        for r in all_rows[1:]:
            if not any(r): continue
            padded = [r[i] if (i < len(r) and r[i] != "") else None for i in range(len(headers))]
            rows.append(padded)
        return headers, rows
    except Exception:
        return [], []

def baca_daftar_kolom_input(in_path, sheet_name=None):
    file_path, sub_sheet = urai_path_excel(in_path)
    target_sheet = sheet_name or sub_sheet
    ext = os.path.splitext(file_path)[1].lower()

    if ext in (".xls", ".xlsx", ".xlsm", ".xlsb") and os.path.isfile(file_path):
        headers, _ = baca_data_excel_python(file_path, target_sheet, hanya_header=True)
        if headers: return headers

    if ext in (".csv", ".txt", ".tab", ".tsv") and os.path.isfile(file_path):
        headers, _ = baca_data_csv_python(file_path, hanya_header=True)
        if headers: return headers

    try:
        target = in_path
        if ext in (".xls", ".xlsx", ".xlsm", ".xlsb") and target_sheet and not sub_sheet:
            target = os.path.join(file_path, target_sheet + "$")
        fields = [
            f.name for f in arcpy.ListFields(target)
            if f.type not in ("OID", "Geometry") and f.name.upper() not in ("SHAPE_LENGTH", "SHAPE_AREA")
        ]
        return fields
    except Exception:
        return []

def cari_indeks_header_persis(headers, target_name):
    t = unicode(target_name or "").strip()
    for i, h in enumerate(headers):
        if unicode(h).strip() == t: return i
    for i, h in enumerate(headers):
        if unicode(h).strip().upper() == t.upper(): return i
    t_clean = re.sub(r'[^A-Za-z0-9_]', '_', t).strip("_").upper()
    for i, h in enumerate(headers):
        h_clean = re.sub(r'[^A-Za-z0-9_]', '_', unicode(h)).strip("_").upper()
        if h_clean == t_clean: return i
    raise ValueError(
        "Kolom '%s' tidak ditemukan di dalam tabel input. Kolom yang tersedia: %s"
        % (target_name, ", ".join(headers))
    )

def siapkan_tabel_gdb_terkunci(in_path, sheet_name, out_gdb, out_table_name,
                               x_field_user, y_field_user, nilai_field_user, surv_field_user, messages):
    file_path, sub_sheet = urai_path_excel(in_path)
    target_sheet = sheet_name or sub_sheet
    ext = os.path.splitext(file_path)[1].lower()
    out_table_full = os.path.join(out_gdb, out_table_name)

    headers, rows = [], []
    if ext in (".xls", ".xlsx", ".xlsm", ".xlsb") and os.path.isfile(file_path):
        messages.addMessage("   Format terdeteksi: Excel (%s) | Sheet: %s" % (ext, target_sheet or "Sheet Pertama"))
        headers, rows = baca_data_excel_python(file_path, target_sheet, hanya_header=False)
    elif ext in (".csv", ".txt", ".tab", ".tsv") and os.path.isfile(file_path):
        messages.addMessage("   Format terdeteksi: Teks/CSV (%s)" % ext)
        headers, rows = baca_data_csv_python(file_path, hanya_header=False)

    if headers and rows:
        idx_x = cari_indeks_header_persis(headers, x_field_user)
        idx_y = cari_indeks_header_persis(headers, y_field_user)
        idx_val = cari_indeks_header_persis(headers, nilai_field_user)
        idx_surv = cari_indeks_header_persis(headers, surv_field_user)

        arcpy.CreateTable_management(out_gdb, out_table_name)
        used_gdb_names = set(["OBJECTID", "FID", "SHAPE", "QC_X_NUM", "QC_Y_NUM", "QC_VAL_NUM", "QC_SURV_TXT"])
        gdb_fields_info = []

        for idx, h_asli in enumerate(headers):
            clean = re.sub(r'[^A-Za-z0-9_]', '_', h_asli).strip("_")
            if not clean: clean = "COL_%d" % (idx + 1)
            if clean[0].isdigit(): clean = "F_" + clean
            clean = clean[:55]

            final_gdb_name = clean
            cnt = 1
            while final_gdb_name.upper() in used_gdb_names:
                suf = "_%d" % cnt
                final_gdb_name = clean[:55 - len(suf)] + suf
                cnt += 1
            used_gdb_names.add(final_gdb_name.upper())

            col_vals = [r[idx] for r in rows if idx < len(r) and r[idx] is not None and unicode(r[idx]).strip() != ""]

            if idx in (idx_x, idx_y):
                arcpy.AddField_management(out_table_full, final_gdb_name, "DOUBLE", field_alias=h_asli[:254])
                gdb_fields_info.append((final_gdb_name, "XY_DOUBLE", 0))
            elif idx == idx_val:
                arcpy.AddField_management(out_table_full, final_gdb_name, "DOUBLE", field_alias=h_asli[:254])
                gdb_fields_info.append((final_gdb_name, "NILAI_DOUBLE", 0))
            else:
                is_numeric = False
                if col_vals:
                    is_numeric = True
                    for v in col_vals:
                        if isinstance(v, (int, long, float)): continue
                        s_v = unicode(v).strip()
                        if len(s_v) > 1 and s_v.startswith("0") and not s_v.startswith("0.") and not s_v.startswith("0,"):
                            is_numeric = False
                            break
                        if not re.match(r'^[-+]?\d+([.,]\d+)?$', s_v):
                            is_numeric = False
                            break
                if is_numeric:
                    arcpy.AddField_management(out_table_full, final_gdb_name, "DOUBLE", field_alias=h_asli[:254])
                    gdb_fields_info.append((final_gdb_name, "DOUBLE", 0))
                else:
                    max_len = 254
                    if col_vals:
                        longest = max(len(unicode(v)) for v in col_vals)
                        max_len = max(254, min(longest + 50, 4000))
                    arcpy.AddField_management(
                        out_table_full, final_gdb_name, "TEXT",
                        field_length=max_len, field_alias=h_asli[:254]
                    )
                    gdb_fields_info.append((final_gdb_name, "TEXT", max_len))

        arcpy.AddField_management(out_table_full, "QC_X_NUM", "DOUBLE")
        arcpy.AddField_management(out_table_full, "QC_Y_NUM", "DOUBLE")
        arcpy.AddField_management(out_table_full, "QC_VAL_NUM", "DOUBLE")
        arcpy.AddField_management(out_table_full, "QC_SURV_TXT", "TEXT", field_length=254)

        insert_cols = [fi[0] for fi in gdb_fields_info] + ["QC_X_NUM", "QC_Y_NUM", "QC_VAL_NUM", "QC_SURV_TXT"]
        with arcpy.da.InsertCursor(out_table_full, insert_cols) as icur:
            for r in rows:
                ins_row = []
                for idx, (_, ftype, flen) in enumerate(gdb_fields_info):
                    v = r[idx] if idx < len(r) else None
                    if v is None: ins_row.append(None)
                    elif ftype == "XY_DOUBLE": ins_row.append(bersihkan_koordinat(v))
                    elif ftype in ("NILAI_DOUBLE", "DOUBLE"): ins_row.append(bersihkan_angka(v))
                    else: ins_row.append(unicode(v)[:flen])

                vx = bersihkan_koordinat(r[idx_x]) if idx_x < len(r) else None
                vy = bersihkan_koordinat(r[idx_y]) if idx_y < len(r) else None
                vval = bersihkan_angka(r[idx_val]) if idx_val < len(r) else None
                vsurv = unicode(r[idx_surv]).strip()[:254] if (idx_surv < len(r) and r[idx_surv] is not None) else None

                ins_row.extend([vx, vy, vval, vsurv])
                icur.insertRow(ins_row)
        return

    target_source = in_path
    if ext in (".xls", ".xlsx", ".xlsm", ".xlsb"):
        try:
            if target_sheet: arcpy.ExcelToTable_conversion(file_path, out_table_full, target_sheet)
            else: arcpy.ExcelToTable_conversion(file_path, out_table_full)
        except Exception:
            if target_sheet and not sub_sheet:
                target_source = os.path.join(file_path, target_sheet + "$")
            arcpy.TableToTable_conversion(target_source, out_gdb, out_table_name)
    else:
        arcpy.TableToTable_conversion(target_source, out_gdb, out_table_name)

    tbl_fields = [f.name for f in arcpy.ListFields(out_table_full)]
    c_x = cari_nama_kolom_ketat(tbl_fields, x_field_user)
    c_y = cari_nama_kolom_ketat(tbl_fields, y_field_user)
    c_val = cari_nama_kolom_ketat(tbl_fields, nilai_field_user)
    c_surv = cari_nama_kolom_ketat(tbl_fields, surv_field_user)

    arcpy.AddField_management(out_table_full, "QC_X_NUM", "DOUBLE")
    arcpy.AddField_management(out_table_full, "QC_Y_NUM", "DOUBLE")
    arcpy.AddField_management(out_table_full, "QC_VAL_NUM", "DOUBLE")
    arcpy.AddField_management(out_table_full, "QC_SURV_TXT", "TEXT", field_length=254)

    with arcpy.da.UpdateCursor(
        out_table_full,
        [c_x, c_y, c_val, c_surv, "QC_X_NUM", "QC_Y_NUM", "QC_VAL_NUM", "QC_SURV_TXT"]
    ) as ucur:
        for row in ucur:
            row[4] = bersihkan_koordinat(row[0])
            row[5] = bersihkan_koordinat(row[1])
            row[6] = bersihkan_angka(row[2])
            row[7] = unicode(row[3]).strip()[:254] if row[3] is not None else None
            ucur.updateRow(row)


def cari_nama_kolom_ketat(daftar_kolom, target, wajib=True):
    if not target: return None
    t_up = unicode(target).strip().upper()
    t_clean = re.sub(r'[^A-Za-z0-9_]', '_', unicode(target).strip()).strip("_").upper()
    if t_clean and t_clean[0].isdigit(): t_clean = "F_" + t_clean
    for k in daftar_kolom:
        if k.upper() == t_up: return k
    for k in daftar_kolom:
        if k.upper() == t_clean: return k
    for k in daftar_kolom:
        if len(k) == 10 and len(t_clean) > 10 and k.upper() == t_clean[:10]: return k
    if wajib:
        raise ValueError("Kolom '%s' tidak ditemukan pada tabel: %s" % (target, ", ".join(daftar_kolom)))
    return None

def normalisasi_id_zona(val):
    if val is None: return None
    if isinstance(val, (int, long)): return unicode(val)
    if isinstance(val, float):
        if math.isnan(val): return None
        if val.is_integer(): return unicode(long(val))
        return unicode(val)
    s = unicode(val).strip()
    if not s or s.upper() in ("NONE", "NULL", "<NULL>"): return None
    if re.match(r'^-?[1-9]\d*\.0+$', s) or s == "0.0":
        s = s.split(".")[0]
    return s

def bersihkan_koordinat(val):
    if val is None: return None
    if isinstance(val, (int, long, float)):
        if math.isnan(float(val)): return None
        return float(val)
    s = unicode(val).strip().replace(" ", "")
    if not s: return None
    if "," in s and "." not in s: s = s.replace(",", ".")
    try: return float(s)
    except ValueError: return None

def simpan_titik_hasil_join(temp_join_fc, output_titik_path):
    is_shp = output_titik_path.lower().endswith(".shp")
    out_dir = os.path.dirname(output_titik_path)
    out_name = os.path.basename(output_titik_path)
    fms = arcpy.FieldMappings()
    used_names = set(["FID", "OBJECTID", "SHAPE"])
    skip_sys = set([
        "JOIN_COUNT", "TARGET_FID", "JOIN_FID", "SHAPE_LENGTH", "SHAPE_AREA",
        "QC_X_NUM", "QC_Y_NUM", "QC_VAL_NUM", "QC_SURV_TXT"
    ])
    for f in arcpy.ListFields(temp_join_fc):
        if f.type in ("OID", "Geometry") or f.name.upper() in skip_sys: continue
        fm = arcpy.FieldMap()
        fm.addInputField(temp_join_fc, f.name)
        out_f = fm.outputField
        clean_name = re.sub(r'[^A-Za-z0-9_]', '_', f.name).strip("_")
        if not clean_name: clean_name = "FIELD"
        if clean_name[0].isdigit(): clean_name = "F_" + clean_name
        max_len = 10 if is_shp else 64
        base_name = clean_name[:max_len]
        final_name = base_name
        cnt = 1
        while final_name.upper() in used_names:
            suf = "_%d" % cnt
            final_name = base_name[:max_len - len(suf)] + suf
            cnt += 1
        used_names.add(final_name.upper())
        out_f.name = final_name
        out_f.aliasName = (f.aliasName or f.name)[:254]
        if out_f.type == "String" and (not out_f.length or out_f.length < 254):
            out_f.length = 254
        fm.outputField = out_f
        fms.addFieldMap(fm)
    arcpy.FeatureClassToFeatureClass_conversion(temp_join_fc, out_dir, out_name, field_mapping=fms)

def bangun_field_mappings_lengkap(fc_titik, fc_zona, id_zona_field, is_shp):
    fms = arcpy.FieldMappings()
    used_names = set(["FID", "OBJECTID", "SHAPE", "JOIN_COUNT", "TARGET_FID", "JOIN_FID"])
    for f in arcpy.ListFields(fc_titik):
        if f.type in ("OID", "Geometry") or f.name.upper() in ("SHAPE_LENGTH", "SHAPE_AREA"): continue
        fm = arcpy.FieldMap()
        fm.addInputField(fc_titik, f.name)
        out_f = fm.outputField
        clean_name = re.sub(r'[^A-Za-z0-9_]', '_', f.name).strip("_")
        if not clean_name: clean_name = "FIELD"
        if clean_name[0].isdigit(): clean_name = "F_" + clean_name
        max_len = 10 if is_shp else 64
        base_name = clean_name[:max_len]
        final_name = base_name
        counter = 1
        while final_name.upper() in used_names:
            suffix = "_%d" % counter
            final_name = base_name[:max_len - len(suffix)] + suffix
            counter += 1
        used_names.add(final_name.upper())
        out_f.name = final_name
        out_f.aliasName = (f.aliasName or f.name)[:254]
        if out_f.type == "String" and (not out_f.length or out_f.length < 254):
            out_f.length = 254
        fm.outputField = out_f
        fms.addFieldMap(fm)
    fm_id = arcpy.FieldMap()
    fm_id.addInputField(fc_zona, id_zona_field)
    fm_id.mergeRule = "First"
    of = fm_id.outputField
    zona_col_name = "ZONA_QC"
    if zona_col_name.upper() in used_names: zona_col_name = "ZONA_ID_QC"
    of.name = zona_col_name
    of.aliasName = "ZONA_QC"
    fm_id.outputField = of
    fms.addFieldMap(fm_id)
    return fms, zona_col_name


# =====================================================================
# FUNGSI SISTEM KOORDINAT & BANTUAN UMUM
# =====================================================================

def sr_tidak_dikenal(sr):
    if sr is None: return True
    name = getattr(sr, "name", "") or ""
    return name == "" or name.lower() == "unknown"

def sr_sama(sr1, sr2):
    if sr_tidak_dikenal(sr1) or sr_tidak_dikenal(sr2): return False
    code1 = getattr(sr1, "factoryCode", 0) or 0
    code2 = getattr(sr2, "factoryCode", 0) or 0
    if code1 != 0 and code1 == code2: return True
    if sr1.name.lower() == sr2.name.lower(): return True
    try: return sr1.exportToString() == sr2.exportToString()
    except Exception: return False

def pilih_transformasi(from_sr, to_sr, dataset=None):
    try:
        extent = arcpy.Describe(dataset).extent if dataset else None
        trans_list = arcpy.ListTransformations(from_sr, to_sr, extent)
        if trans_list: return trans_list[0]
    except Exception:
        pass
    return ""

def cek_rentang_xy(table, x_field, y_field, sr, messages):
    xs, ys = [], []
    with arcpy.da.SearchCursor(table, [x_field, y_field]) as cur:
        for x, y in cur:
            vx = bersihkan_koordinat(x)
            vy = bersihkan_koordinat(y)
            if vx is not None and vy is not None:
                xs.append(vx)
                ys.append(vy)
    if not xs:
        arcpy.AddError("Tidak ada pasangan koordinat X dan Y yang valid (angka) pada tabel input.")
        raise ValueError("Koordinat X/Y kosong atau bukan angka.")
    min_x, max_x = min(xs), max(xs)
    min_y, max_y = min(ys), max(ys)
    if getattr(sr, "type", "") == "Geographic":
        if min_x < -180 or max_x > 180 or min_y < -90 or max_y > 90:
            messages.addWarningMessage(
                "   PERINGATAN: Sistem koordinat dipilih Geographic (derajat), "
                "tetapi rentang X (%.2f s/d %.2f) atau Y (%.2f s/d %.2f) di luar batas derajat! "
                "Periksa apakah koordinat sebenarnya UTM/TM3."
                % (min_x, max_x, min_y, max_y)
            )
        elif min_y > 90 or (min_x >= -15 and max_x <= 10 and min_y >= 90):
            messages.addWarningMessage("   PERINGATAN: Kolom X dan Y kemungkinan tertukar (Latitude vs Longitude).")
    else:
        if -180 <= min_x <= 180 and -90 <= min_y <= 90:
            messages.addWarningMessage(
                "   PERINGATAN: Sistem koordinat dipilih Projected (meter), "
                "tetapi angka X/Y terlihat seperti derajat desimal (Geographic)."
            )

def cek_overlap(fc_titik, fc_zona, messages):
    try:
        ext_pt = arcpy.Describe(fc_titik).extent
        ext_zn = arcpy.Describe(fc_zona).extent
        overlap = not (
            ext_pt.XMax < ext_zn.XMin or
            ext_pt.XMin > ext_zn.XMax or
            ext_pt.YMax < ext_zn.YMin or
            ext_pt.YMin > ext_zn.YMax
        )
        if not overlap:
            messages.addWarningMessage(
                "   PERINGATAN: Extent titik survei dan poligon zona TIDAK saling beririsan! "
                "Periksa kesesuaian sistem koordinat atau kemungkinan kolom X dan Y tertukar."
            )
    except Exception: pass

def bersihkan_angka(val):
    if val is None: return None
    if isinstance(val, (int, long, float)):
        if math.isnan(float(val)): return None
        return float(val)
    s = unicode(val).strip()
    if not s or s in ("-", "None", "NULL", "null", "#N/A", "N/A"): return None
    try:
        if "e" in s.lower() and "," not in s: return float(s)
    except Exception: pass
    s = re.sub(r'(?i)(rp|idr|per\s*m2|/m2|[a-z\s])', '', s)
    if not s: return None
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."): s = s.replace(".", "").replace(",", ".")
        else: s = s.replace(",", "")
    elif "," in s:
        parts = s.split(",")
        if len(parts) > 2: s = s.replace(",", "")
        elif len(parts) == 2:
            if len(parts[1]) == 3 and parts[1] == "000" and len(parts[0]) >= 1: s = s.replace(",", "")
            else: s = s.replace(",", ".")
    elif "." in s:
        parts = s.split(".")
        if len(parts) > 2: s = s.replace(".", "")
        elif len(parts) == 2:
            if len(parts[1]) == 3 and parts[1] == "000" and len(parts[0]) >= 2: s = s.replace(".", "")
    try: return float(s)
    except ValueError: return None


# =====================================================================
# FUNGSI PENDUKUNG SUM-ZNT (Helper EkstrakZonaExcel)
# =====================================================================
def to_text(v):
    if v is None:
        return text_type("")
    if isinstance(v, text_type):
        return v
    try:
        return text_type(v)
    except Exception:
        return text_type(str(v))

def parse_number(v):
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, numbers.Number):
        f = float(v)
        return None if f != f else f
    s = re.sub(r"[^0-9,.\-]", "", to_text(v))
    if not re.search(r"[0-9]", s):
        return None
    neg = s.startswith("-")
    s = s.replace("-", "")
    if "." in s and "," in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s or "." in s:
        sep = "," if "," in s else "."
        parts = s.split(sep)
        if len(parts) > 2:
            s = "".join(parts)
        else:
            head, tail = parts
            if len(tail) == 3 and 1 <= len(head) <= 3 and head != "0":
                s = head + tail
            else:
                s = head + "." + tail
    try:
        val = float(s)
    except ValueError:
        return None
    return -val if neg else val

def natural_key(s):
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", s)]

def resolve_field(dataset, wanted):
    names = [f.name for f in arcpy.ListFields(dataset)]
    w = to_text(wanted).lower()
    for n in names:
        if n.lower() == w:
            return n
    short = w.split(".")[-1]
    for n in names:
        if n.split(".")[-1].lower() == short:
            return n
    raise Exception("Field '{0}' tidak ditemukan. Field yang tersedia: {1}".format(
        wanted, ", ".join(names)))

def _sum_make_field_mappings(target_fc, target_fields, join_fc, join_field, out_name):
    fms = arcpy.FieldMappings()
    for fname in target_fields:
        fm = arcpy.FieldMap()
        fm.addInputField(target_fc, fname)
        fms.addFieldMap(fm)
    fm = arcpy.FieldMap()
    fm.addInputField(join_fc, join_field)
    out_field = fm.outputField
    out_field.name = out_name
    out_field.aliasName = out_name
    fm.outputField = out_field
    fm.mergeRule = "First"
    fms.addFieldMap(fm)
    return fms

def _sum_datum_name(sr):
    try: return sr.GCS.datumName
    except Exception: pass
    try: return sr.datumName
    except Exception: return ""

def _sum_pick_transformation(src_dataset, sr_target):
    try:
        desc = arcpy.Describe(src_dataset)
        d1 = _sum_datum_name(desc.spatialReference)
        d2 = _sum_datum_name(sr_target)
        if not d1 or not d2 or d1 == d2: return ""
        lst = arcpy.ListTransformations(desc.spatialReference, sr_target, desc.extent)
        return lst[0] if lst else ""
    except Exception: return ""

def _sum_slim_copy(src, out_fc, field_pairs, sr_target):
    fms = arcpy.FieldMappings()
    for src_field, out_name in field_pairs:
        fm = arcpy.FieldMap()
        fm.addInputField(src, src_field)
        of = fm.outputField
        of.name = out_name
        of.aliasName = out_name
        fm.outputField = of
        fm.mergeRule = "First"
        fms.addFieldMap(fm)
    tf = _sum_pick_transformation(src, sr_target)
    try: old_tf = arcpy.env.geographicTransformations
    except Exception: old_tf = ""
    try:
        if tf: arcpy.env.geographicTransformations = tf
        out_path, out_name = os.path.split(out_fc)
        try: arcpy.FeatureClassToFeatureClass_conversion(src, out_path, out_name, "", fms)
        except AttributeError: arcpy.conversion.ExportFeatures(src, out_fc, "", "NOT_USE_ALIAS", fms)
    finally:
        try: arcpy.env.geographicTransformations = old_tf if old_tf else ""
        except Exception: pass
    try: arcpy.RepairGeometry_management(out_fc)
    except Exception: pass
    return out_fc

def _sum_area_m2(shape, geodesic):
    try: return shape.getArea("GEODESIC" if geodesic else "PLANAR", "SQUAREMETERS")
    except Exception: return shape.area

def _sum_new_stats():
    return {"total": 0, "zona_kosong": 0, "harga_kosong": 0,
            "harga_bukan_angka": 0, "fragmen_kecil": 0, "valid": 0}

def _sum_clean_row(zona_raw, harga_raw, stats):
    zona = to_text(zona_raw).strip()
    if not zona:
        stats["zona_kosong"] += 1
        return None
    if harga_raw is None or not to_text(harga_raw).strip():
        stats["harga_kosong"] += 1
        return None
    harga = parse_number(harga_raw)
    if harga is None:
        stats["harga_bukan_angka"] += 1
        return None
    return zona, harga

def _sum_aggregate_by_kecamatan(records, desa_lookup, enable_desa):
    by_kec = {}
    for kec, zona, harga, desa in records:
        by_kec.setdefault(kec, []).append((harga, zona, desa))
    def pick(kec, recs, value):
        zonas = set()
        desas = set()
        for harga, zona, desa in recs:
            if harga != value: continue
            zonas.add(zona)
            if not enable_desa: continue
            if desa_lookup is not None:
                desas |= desa_lookup.get((kec, zona, harga), set())
            elif desa and desa != "-":
                desas.add(desa)
        if enable_desa and not desas:
            desas.add(TIDAK_TERIDENTIFIKASI)
        return zonas, desas
    def join(items):
        return ", ".join(sorted(items, key=natural_key))[:1000]
    rows = []
    for kec in sorted(by_kec.keys()):
        recs = by_kec[kec]
        lo = min(r[0] for r in recs)
        hi = max(r[0] for r in recs)
        z_lo, d_lo = pick(kec, recs, lo)
        z_hi, d_hi = pick(kec, recs, hi)
        rows.append([kec[:150], lo, join(z_lo), join(d_lo), hi, join(z_hi), join(d_hi)])
    return rows