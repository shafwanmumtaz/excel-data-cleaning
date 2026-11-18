# Pembersihan & Standarisasi Data Penjualan Karyawan (Q1–Q4) serta Analisis Statistika di Microsoft Excel

👤 **Penyusun:** Shafwan Mumtaz  
📊 **Tools:** Microsoft Excel (`TRIM`, `PROPER`, `TEXT`, `IF`, `COUNTIF`, `AVERAGE`, `MAX`, `MIN`, Perhitungan Penyebaran Data & Statistika Deskriptif)

---

## 1. Gambaran Umum Proyek

Repositori ini mendokumentasikan pekerjaan pengolahan, pembersihan (*data cleaning*), serta perhitungan statistik menggunakan **Microsoft Excel**, meliputi:

1. **Pembersihan & Standarisasi Data Penjualan Karyawan Kuartalan (Q1–Q4):**
   - Mengatasi permasalahan spasi berlebih dan penulisan nama karyawan/klien yang tidak seragam menggunakan kombinasi fungsi `=PROPER(TRIM(...))`.
   - Menyeragamkan format tanggal transaksi menggunakan `=TEXT(Tanggal, "YYYY-MM-DD")`.
   - Mengisi status capaian kuartalan secara otomatis menggunakan fungsi logika kondisional `=IF(...)`.
   - Menghitung ringkasan performa menggunakan `COUNTIF`, `AVERAGE`, `MAX`, dan `MIN`.

2. **Perhitungan Statistika Deskriptif & Penyebaran Data (`latihan_penyebaran_data.xlsx` & `latihan_statistika.xlsx`):**
   - Pengolahan tabel distribusi frekuensi berkelompok (nilai tengah $X_i$, tepi bawah $tb$, tepi atas $ba$).
   - Perhitungan **Rata-rata ($ar{x}$)**, **Variansi ($s^2$)**, dan **Simpangan Baku ($s = \sqrt{s^2}$)**.
   - Perhitungan ukuran letak data: **Kuartil ($Q_1, Q_2, Q_3$)**, **Desil**, dan **Persentil**.

3. **Pengolahan Nilai & Grade Akademik (`Shafwan_Mumtaz_Aplikasi_Komputer.xlsx`):**
   - Rekapitulasi nilai Tugas, UTS, dan UAS untuk menentukan Nilai Akhir, Grade huruf, serta status kelulusan menggunakan rumus Excel.

---

## 2. Daftar File dalam Repositori

```text
├── data/
│   ├── employee_sales_q1_q4_raw.csv           # Contoh data penjualan mentah sebelum dibersihkan
│   └── employee_sales_q1_q4_cleaned.csv       # Data penjualan bersih hasil standarisasi rumus Excel
├── latihan_penyebaran_data.xlsx               # Workbook Excel perhitungan Simpangan Baku, Kuartil, Desil, Persentil
├── latihan_statistika.xlsx                    # Workbook Excel distribusi frekuensi & rata-rata berkelompok
├── Shafwan_Mumtaz_Aplikasi_Komputer.xlsx      # Workbook Excel pengolahan nilai & fungsi logika
└── README.md                                  # Dokumentasi proyek
```
