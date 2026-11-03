# Aplikasi Stress Test Risiko (Permintaan Data OJK)

Aplikasi GUI (Streamlit) untuk membuat file **"Stress Test Risiko &lt;Bulan&gt; &lt;Tahun&gt;.xlsx"**:
risiko pasar, kredit, dan likuiditas, dengan format **sama persis** seperti file
*Permintaan OJK Datastress test risiko pasar kredit likuiditas* (sheet, warna, font,
rumus, lebar kolom, sel kuning, dan komentar tetap sama).

---

## 1. Cara menjalankan

1. Pasang **Python 3.10 atau lebih baru** (centang *Add Python to PATH* saat instalasi).
2. Klik ganda **`Jalankan_Aplikasi.bat`**. Library akan terpasang otomatis, lalu aplikasi terbuka di browser.
   Cara manual lewat CMD di folder ini:
   ```
   pip install -r requirements.txt
   streamlit run app_stress_test.py
   ```
3. File hasil tersimpan di folder **`hasil/`** dan juga bisa diunduh dari tab 7.

## 2. Alur kerja setiap bulan

| Tab | Yang dilakukan | Sumber data | Masuk ke Excel |
|---|---|---|---|
| Samping | Pilih **bulan & tahun laporan** (otomatis dari Rekap) | – | Nama file, tanggal B2, teks "(30 Sept 2026)", komentar |
| 1 · RAW Investasi | Upload **Rekap Investasi** (.xlsx). Tabel dicari otomatis berdasarkan nama kolom | DISK | Sheet `RAW INVESTASI` (seluruh kolom, urutan template) |
| 2 · IHSG, Yield & Kurs | Isi IHSG existing, IHSG severe, yield, kurs. **IHSG Escalated** diatur dengan slider yang hanya bisa bergerak di antara IHSG Severe dan IHSG Baseline. Untuk Escalated/Severe, pilih salah satu: isi **Yield skenario** (B dihitung = Yield − Yield Awal), atau centang **Isi Peningkatan Suku Bunga (B) manual** (Yield skenario = Yield Awal + B) | Investing.com, ketentuan OJK, tim MR | `Hasil Stress Testing` C4:F6, `1-Risiko Pasar` baris 18, `3-Risiko Likuiditas` D10 |
| 3 · Kalkulator Baseline | Kotak input kalkulator IHSG & Yield baseline, proyeksi Desember, existing, dan Aset Neto. Bisa **isi otomatis** dengan upload Budget & Realisasi (opsional) | Budget & Realisasi (ASUMSI HU, ASUMSI NERACA) | `Skenario Baseline`, `Hasil` C10:C13, L4:L7, L12 |
| 4 · Komposisi RD | Default Schroder Dana Terpadu II & Trimegah Balanced Absolute Strategy Kelas A, bisa diganti/ditambah. **Ambil FFS terbaru** mengunduh fund fact sheet bulan laporan dan membaca porsi Saham/Obligasi/Pasar Uang. Ada kotak input manual **SPI Reksa Dana per Jenis** | Fund fact sheet, laporan DISK *SPI Reksadana per Jenis* | `Komposisi RD` |
| 5 · Risiko Kredit | Eksposur per grup dihitung dari RAW. Isi/ubah **rating grup** dan **rating MI** bila ada yang baru atau berubah | RAW + Display Monitoring Risiko Investasi (rating) | `2-Risiko Kredit`, `2-CHECK (1)`, `2-Check (2)` |
| 6 · Risiko Likuiditas | Eksposur (keterlambatan iuran + manfaat pensiun) & tenor | Tim MR | `3-Risiko Likuiditas` |
| 7 · Ringkasan | Cek kelengkapan, lalu **Hitung & Buat Excel** | – | Semua sheet |

Semua input disimpan ke `config/input_terakhir.json` setiap kali Excel dibuat. Jadi bulan
berikutnya cukup mengganti angka yang berubah.

### Rincian angka dari Budget & Realisasi (isi otomatis)

| Isian | Sel sumber (contoh September 2026) |
|---|---|
| Existing: Total Aset Investasi / SPI / Laba Usaha / Delta SPI | ASUMSI HU: *RATA-RATA NILAI WAJAR* Total, *SPI* Total, *PENDAPATAN INVESTASI BERSIH YTD* Total, *DELTA SPI YTD* Total pada kolom bulan laporan (K330, K246, K204, K267) ÷ 1.000 |
| Proyeksi Desember: keempat angka yang sama | Kolom Desember (N330, N246, N204, N267) ÷ 1.000 |
| SPI Saham 31 Des / bulan laporan (kalkulator IHSG) | ASUMSI HU › SPI › Saham (N236 / K236), dalam juta |
| Delta SPI selain saham (kalkulator yield) | SPI (Total − Saham − Penyertaan Langsung − Tanah Bangunan): Desember dikurangi bulan laporan, ÷ 1.000, dibulatkan 3 desimal (0,471) |
| Aset Neto | ASUMSI NERACA › Total Aset Neto bulan laporan × 1.000.000 |

Kalkulator yield tidak lagi memerlukan *Goal Seek*: aplikasi menghitung langsung
`C5 = Delta SPI ÷ Nilai Wajar + PV Existing`. Hasilnya sama dengan Goal Seek
(September 2026: 7,1036%).

## 3. Fund fact sheet

* Daftar link FFS tersimpan di `config/katalog_reksadana.json`. Link boleh memakai penanda
  `{bulan} {yyyy}` (bulan data, mis. *September 2026*) atau `{pub_yyyy}-{pub_m}` (bulan terbit = bulan laporan + 1).
* Aplikasi mencoba link satu per satu dan memilih FFS yang **tanggal datanya sama dengan bulan laporan**.
  Bila yang ditemukan masih FFS bulan sebelumnya, muncul peringatan.
* Bila semua link gagal (situs berubah atau tidak ada internet), buka **Link fund fact sheet**,
  unduh PDF-nya, lalu upload di **…atau upload PDF FFS**, atau ketik persentasenya.
  Link baru yang Anda tempel di kolom *Tambah link FFS sendiri* otomatis disimpan ke katalog.

## 4. Hal yang perlu diketahui

* **Kesamaan dengan file asli.** Membuat ulang September 2026 dari Rekap Investasi + Budget
  menghasilkan file yang identik dengan file asli di sheet Hasil, Skenario Baseline, Komposisi RD,
  Risiko Pasar Escalated/Severe, dan Risiko Likuiditas (nilai, rumus, dan gaya).
  Perbedaan yang disengaja:
  * Pivot di `2-CHECK (1)` dan `CHECK` diganti tabel bertampilan sama yang berisi rumus SUMIF,
    karena pivot di template tidak ikut ter-refresh. Rumus GETPIVOTDATA di E3:E6 diganti
    penjumlahan sel yang setara.
  * Nomor urut di `2-Risiko Kredit` dibuat berurutan (di file asli ada nomor "25" di baris ke-12).
  * Sheet tersembunyi (`1-Risiko Pasar Baseline/Existing`, `Rating`, `CHECK`) ikut diperbarui ke bulan laporan.
  * Rumus L9, M9, L11, M11 (sheet *1-Risiko Pasar*) dan `Hasil Stress Testing`!C40 memakai versi terbaru.
* File hasil sudah berisi angka hasil hitung, jadi angka tetap tampil walau dibuka di Protected View
  atau pratinjau email. Excel tetap menghitung ulang saat file dibuka.
* Eksposur kredit ditulis dalam miliar dengan 6 desimal (dipotong per Rp1.000), sama dengan file asli.

## 5. Isi folder

```
app_stress_test.py        GUI Streamlit
stress_core.py            Baca Rekap & Budget, hitung baseline, tulis Excel dari template
xl_engine.py              Penghitung rumus Excel (untuk pratinjau & nilai tersimpan)
ffs_reader.py             Unduh & baca fund fact sheet
template/                 Template (file Stress Test September 2026)
config/                   Katalog FFS, master rating, input terakhir
hasil/                    File Excel hasil
Jalankan_Aplikasi.bat     Klik ganda untuk menjalankan (Windows)
```
