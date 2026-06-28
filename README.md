# Generator Laporan Absensi Otomatis

Aplikasi ini adalah solusi otomatisasi berbasis Python yang dibuat untuk mempermudah konversi data mentah absensi (dari CSV mesin absensi) ke dalam format **Laporan Excel Resmi**. 

Aplikasi ini akan secara otomatis:
- Menghitung durasi kerja (Jam Pulang - Jam Masuk).
- Mendeteksi hari libur, akhir pekan (Sabtu/Minggu), dan status cuti.
- Mengatur tata letak, warna huruf (teks merah untuk hari libur), dan presisi baris tanpa merusak format template Excel bawaan.
- Tersedia dalam dua mode: **Terminal/CLI** (interaktif) dan **GUI** (Visual Desktop).

---

## 🛠️ Persyaratan Sistem (Prerequisites)
Sebelum menggunakan aplikasi ini, pastikan komputer/laptop Anda telah terinstal:
- **Python 3.8** atau versi lebih baru. 

---

## 🚀 Cara Instalasi

Ikuti langkah-langkah di bawah ini untuk mengatur *environment* agar aplikasi bisa berjalan dengan baik:

1. **Buka Terminal / Command Prompt**
2. **Masuk ke folder proyek ini** (sesuaikan dengan lokasi Anda menyimpan folder ini):
   ```bash
   cd /path/ke/folder/laporan-absensi
   ```
3. **Buat Virtual Environment (Sangat disarankan)**
   Untuk menjaga agar *library* tidak bentrok dengan aplikasi lain:
   ```bash
   python3 -m venv .venv
   ```
4. **Aktifkan Virtual Environment**
   - **Mac/Linux:**
     ```bash
     source .venv/bin/activate
     ```
   - **Windows:**
     ```cmd
     .venv\Scripts\activate
     ```
5. **Instal Dependencies (Library yang dibutuhkan)**
   ```bash
   pip install -r requirements.txt
   ```
   *(Ini akan menginstal `pandas`, `openpyxl`, dan `customtkinter`)*

---

## 💻 Cara Penggunaan

Anda dapat memilih satu dari dua cara untuk menjalankan aplikasi ini:

### Opsi A: Versi Visual / GUI (Rekomendasi)
Tampilan desktop modern yang ramah pengguna. Sangat cocok jika teman Anda tidak terbiasa dengan layar kode terminal.
```bash
python gui_app.py
```
**Langkah Penggunaan GUI:**
1. Klik tombol **Browse** untuk memilih file Data Mentah (`.csv`).
2. Setelah file dipilih, menu *dropdown* **Nama Pegawai** akan otomatis terisi dengan daftar nama. Klik nama yang ingin dicetak laporannya.
3. Anda bisa mengedit nilai **OPD**, **Nama Projek**, dan **Role Pegawai** (Nilai bawaan sudah disediakan dan siap pakai).
4. Klik tombol besar **Generate Laporan Excel**.
5. Tunggu *pop-up* "Berhasil" muncul!

### Opsi B: Versi Terminal / CLI (Cepat & Ringan)
Menggunakan terminal dengan sistem tanya-jawab (*wizard*) interaktif. Sangat cepat.
```bash
python app.py
```
**Langkah Penggunaan Terminal:**
1. Anda akan diminta memasukkan nama file CSV (cukup tekan `Enter` untuk menggunakan file *default*).
2. Sistem akan menampilkan daftar nama pegawai yang ditemukan. Ketik **angka urut** dari pegawai yang dipilih lalu tekan `Enter`.
3. Sistem akan mengonfirmasi isian **OPD**, **Nama Projek**, dan **Role**. (Cukup tekan `Enter` berkali-kali jika Anda ingin memakai *default*-nya).
4. Sistem memproses file dan memberikan laporan sukses.

---

## 📁 Struktur File Penting

- `Laporan Absensi Maret 2026.xlsx` : **Master Template Excel**. Jangan ubah posisi *header* atau kolom tabel di file ini karena skrip menjadikannya patokan utama untuk bekerja.
- `ekspor_csv Mei 26.csv` : Contoh data mentah yang dikeluarkan dari mesin absensi. Anda bisa menggunakan file csv lain asal strukturnya sama.
- `gui_app.py` & `app.py` : Skrip eksekutor aplikasi (GUI dan Terminal).
- `requirements.txt` : Daftar pustaka (*library*) wajib untuk Python.

---

## 📈 Hasil Output
File yang berhasil digenerate akan muncul di folder yang sama dengan format penamaan otomatis menggunakan ekstensi *.xlsx*, contoh:
`Laporan_Absensi_Brahmanda_20260628_224500.xlsx`

**Kelebihan Output yang Dihasilkan:**
- Mengamankan blok Tanda Tangan di bawah (selalu diberikan jarak presisi tepat 2 baris kosong).
- Tanggal pengesahan tanda tangan akan otomatis mencetak tanggal hari ini di mana Anda men-generate laporan tersebut.
- *Text* berwarna merah menyala otomatis untuk hari **Sabtu**, **Minggu**, dan status **Tidak Hadir / Cuti** pada seluruh kolom terkait.
