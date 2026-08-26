# Pengantar Komputer, Excel, dan VBA untuk Teknik Sipil

**Bahasa:** Indonesia - [English version](raw/utama-en/README.md)

Paket materi Markdown untuk Program Sarjana Teknik Sipil, Departemen Teknik Sipil dan Lingkungan, Universitas Gadjah Mada. Excel dan VBA menjadi lingkungan utama, dengan Python sebagai pembanding singkat dan VBA AutoCAD sebagai demonstrasi opsional.

Versi ini adalah gabungan dua paket modul yang mengajarkan pokok yang sama dari dua sudut pandang: satu bergerak dari perangkat keras dan Excel menuju kode, satu lagi bergerak dari kasus teknik menuju kode. Naskah aslinya tersimpan di [`raw/`](raw/README.md).

## Slide dan PDF

Slide PowerPoint dibuat dengan template UGM dari `_Slide-Template-UGM` dan disimpan bersama bahan mentah utama.

| Bahasa | Slide PowerPoint | PDF Pertemuan 1 | Builder |
|---|---|---|---|
| Indonesia | [`raw/utama-id/00-Slide-Kuliah/`](raw/utama-id/00-Slide-Kuliah/) | [`01-pengantar-komputer-os.pdf`](raw/utama-id/01-pengantar-komputer-os.pdf) | [`build_powerpoints.py`](raw/utama-id/build_powerpoints.py) |
| English | [`raw/utama-en/00-Slide-Kuliah/`](raw/utama-en/00-Slide-Kuliah/) | [`01-computers-and-operating-systems.pdf`](raw/utama-en/01-computers-and-operating-systems.pdf) | [`build_powerpoints.py`](raw/utama-en/build_powerpoints.py) |

## Desain pembelajaran

Setiap modul dirancang untuk **2 SKS atau 100 menit**. Alur pembelajaran dibuat sederhana:

```text
lihat kasus -> prediksi -> pelajari konsep -> demonstrasi -> praktik
            -> uji hasil -> verifikasi manual -> jelaskan dengan kata-kata sendiri
```

Setiap pertemuan menghasilkan satu artefak kecil yang dapat didemonstrasikan dan ditutup dengan tiga soal berdampak tinggi:

1. **Prediksi:** mahasiswa menebak hasil sebelum menjalankan program.
2. **Praktik/perbaikan:** mahasiswa membuat atau memperbaiki artefak.
3. **Jelaskan:** mahasiswa menerangkan alasan, bukan menghafal sintaks.

Satu prinsip berlaku di seluruh modul: **program yang menghasilkan angka belum tentu benar.** Setiap hasil disertai satuan, asumsi, dan satu cara memeriksanya.

## Daftar modul

| Modul | Materi | Produk minimum |
|---:|---|---|
| 1 | [Pengantar komputer, OS, dan komputasi teknik sipil](01-pengantar-komputer-dan-komputasi.md) | peta kerja komputer, workbook `.xlsm`, dan macro pertama |
| 2 | [Dasar Excel, fungsi, dan disiplin satuan](02-excel-fungsi-dan-satuan.md) | lembar hitungan dengan acuan, fungsi, dan satuan yang benar |
| 3 | [Keputusan dan pengujian batas](03-keputusan-dan-pengujian-batas.md) | klasifikasi yang teruji pada setiap nilai batas |
| 4 | [Algoritma: representasi, penelusuran, dan verifikasi](04-algoritma-dan-verifikasi.md) | IPO, pseudocode, flowchart, tabel jejak, dan hitungan manual |
| 5 | [VBA: macro linier, percabangan, dan perulangan](05-vba-macro-percabangan-perulangan.md) | macro yang memproses banyak baris dengan validasi |
| 6 | [Modularitas, struktur data, debugging, dan pengujian](06-modularitas-data-dan-pengujian.md) | fungsi terdokumentasi beserta pengujian otomatisnya |
| 7 | [UserForm, otomasi terpadu, dan praktik integratif](07-userform-otomasi-dan-praktik-integratif.md) | evaluator alternatif saluran berbasis form dan terverifikasi |

## Pola waktu 100 menit

| Tahap | Durasi | Aktivitas |
|---|---:|---|
| Orientasi | 5-10 menit | kasus, tujuan, dan prediksi awal |
| Konsep inti | 15-25 menit | penjelasan singkat dan contoh konkret |
| Demonstrasi | 15-20 menit | dosen memperlihatkan proses lengkap |
| Praktik | 30-40 menit | mahasiswa membuat dan mengubah contoh |
| Kuis dan refleksi | 10-15 menit | tiga soal, pemeriksaan, dan penjelasan |

Setiap modul memuat tabel **Alur 100 menit** dan **Keluaran minimum**. Bagian bertanda **pengayaan/PR** tidak wajib selesai selama tatap muka. Dosen dapat menyediakan workbook dengan judul kolom dan data awal agar waktu kelas dipakai untuk memahami algoritma dan kode, bukan mengetik tabel.

## Perangkat

- Microsoft Excel desktop dengan VBA;
- workbook disimpan sebagai `.xlsm`;
- opsional: browser dengan Google Colab, atau Python 3 dengan Jupyter, untuk pembanding singkat;
- opsional: AutoCAD dengan lingkungan VBA untuk satu demo; dan
- browser untuk membaca materi Markdown melalui GitHub.

> **Keamanan macro.** Aktifkan macro hanya dari sumber yang dipercaya. Jangan mengaktifkan macro pada lampiran yang tidak dikenal. Simpan salinan data sebelum menjalankan macro yang mengubah banyak sel atau objek gambar.

## Konvensi

- Semua module VBA dimulai dengan `Option Explicit`.
- Nama variabel menyertakan makna dan satuan, misalnya `panjang_m` atau `debit_m3s`.
- Objek Excel dirujuk secara eksplisit, misalnya `ThisWorkbook.Worksheets("Nama").Range("B2")`, bukan mengandalkan lembar yang kebetulan aktif.
- Contoh formula Excel memakai nama fungsi berbahasa Inggris. Pemisah argumen dapat berupa koma atau titik koma, mengikuti pengaturan Excel.
- **F4** mengganti bentuk acuan relatif/absolut saat mengedit formula. **F5** membuka dialog *Go To*. **Alt+F11** membuka Visual Basic Editor. **Ctrl+G** membuka Immediate Window.
- Fungsi hitungan mengembalikan `-1` untuk input yang tidak valid, sehingga pemanggil wajib memeriksa hasil sebelum memakainya.
- Hasil program selalu diperiksa dengan kasus sederhana atau hitungan manual.

## Batas materi teknik

Angka, batas kategori, dan koefisien pada modul ini adalah **data latihan algoritma**, bukan kriteria desain. Persamaan seperti Manning dipakai untuk melatih penerjemahan rumus menjadi kode. Penggunaan untuk desain memerlukan penetapan parameter, kondisi lapangan, dan standar yang berlaku.

## Skema asesmen singkat

| Aspek | Bobot |
|---|---:|
| Artefak berfungsi dan hasil benar | 30% |
| Algoritma, satuan, dan validasi | 25% |
| Pengujian dan verifikasi manual | 15% |
| Tiga soal kuis | 15% |
| Kemampuan mendemonstrasikan dan menjelaskan | 15% |

## Etika penggunaan AI

AI boleh dipakai untuk menjelaskan pesan galat atau menawarkan alternatif kode jika dosen mengizinkan. Mahasiswa tetap wajib mampu menjelaskan setiap baris penting, menyebutkan bantuan yang dipakai, menguji program dengan datanya sendiri, dan bertanggung jawab atas hasil akhir.

---

Mulai dari [Modul 1](01-pengantar-komputer-dan-komputasi.md).
