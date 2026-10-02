# Toolkit-ZNT

Toolkit ini mencakup 3 (tiga) modul utama pendukung manajemen dan validasi Zona Nilai Tanah:

* **Generate Zona Tepi Jalan:** Otomasi deliniasi zona tepi jalan menggunakan metode transek tegak lurus. Alat ini mendeteksi persil terdekat dalam radius pencarian, membuat *buffer* pada sisi tanpa persil, serta mengisi celah (*fill holes*) untuk menghasilkan satu poligon zona yang kohesif dan utuh.


* **Bagi Zona Surveyor:** Membagi poligon zona secara proporsional kepada sejumlah tim surveyor menggunakan algoritma *Recursive Spatial Bisection*. Memastikan beban kerja terdistribusi sama rata dengan area survei yang mengelompok secara spasial.
* **QC Survei Zona:** Memvalidasi kualitas data titik survei dari CSV yang diintegrasikan secara spasial ke poligon zona. Fitur ini secara otomatis menghitung nilai rata-rata dan *Relative Standard Deviation* (RSD), menentukan status kelulusan QC berdasarkan batas toleransi skala proyek (misal: 1:10.000), serta merekapitulasi data surveyor per zona.



**Panduan Download, Instalasi & Pembaruan:**

1. Unduh file `InstallerToolkitZNT.exe` pada halaman [release](https://github.com/ekabudiku/Toolkit-ZNT/releases/tag/ZNT).
2. Letakkan file tersebut di PC/Laptop Anda (direkomendasikan di *Desktop* agar mudah diakses).
3. Klik ganda (2x) aplikasi tersebut untuk mengunduh dan memasang Toolkit secara otomatis ke folder `Documents\Toolkit ZNT`.
4. Buka ArcGIS, klik kanan pada ArcToolbox -> **Add Toolbox** -> Arahkan ke file `ToolkitZNT.pyt` di folder tersebut.
5. Untuk pembaruan di masa mendatang, Anda cukup menutup ArcGIS dan menjalankan kembali `.exe` tersebut.
