# Algoritma dan Pemrograman Komputer untuk Teknik Sipil

Modul pembelajaran berbasis **VBA di Microsoft Excel** untuk mahasiswa Program Sarjana Teknik Sipil, Departemen Teknik Sipil dan Lingkungan, Universitas Gadjah Mada. Python digunakan secara minor sebagai pembanding agar mahasiswa mengenal ekosistem komputasi yang lebih luas.

## Mengapa VBA dan Excel?

Mahasiswa teknik sipil sudah akrab dengan tabel, sel, dan rumus Excel. VBA memperluas Excel dari kalkulator lembar kerja menjadi alat otomasi: membaca banyak data, mengulang perhitungan, memeriksa input, serta menulis hasil secara konsisten. Konsep algoritma yang dipelajari tetap dapat dipindahkan ke Python, MATLAB, R, atau bahasa lain.

## Capaian pembelajaran paket modul

Setelah menyelesaikan tujuh modul, mahasiswa mampu:

1. menguraikan masalah teknik sipil menjadi input–proses–output;
2. menyusun algoritma, flowchart, dan pseudocode;
3. menerapkan variabel, tipe data, percabangan, perulangan, dan struktur data;
4. menulis fungsi dan prosedur VBA yang modular dan terdokumentasi;
5. melakukan validasi input, debugging, penanganan galat, dan pengujian;
6. mengotomasi perhitungan teknik sipil sederhana di Excel; dan
7. memverifikasi hasil program terhadap hitungan manual atau kasus acuan.

## Daftar modul

| Pertemuan | Modul | Produk belajar utama |
|---:|---|---|
| 1 | [Pengantar komputasi dalam teknik sipil](01-pengantar-komputasi.md) | Makro pertama dan peta input–proses–output |
| 2 | [Representasi masalah](02-representasi-masalah.md) | Algoritma, flowchart, pseudocode, dan kalkulator debit |
| 3 | [Struktur kontrol](03-struktur-kontrol.md) | Klasifikasi saluran dan validasi input |
| 4 | [Struktur data](04-struktur-data.md) | Pengolahan deret hujan dan matriks data |
| 5 | [Fungsi, debugging, dan pengujian](05-fungsi-debugging-pengujian.md) | Fungsi teruji untuk hitungan teknik |
| 6 | [Otomasi perhitungan teknik sipil](06-otomasi-perhitungan-teknik-sipil.md) | Konversi satuan, interpolasi, dan hitungan berulang |
| 7 | [Praktik integratif pra-UTS](07-praktik-integratif-pra-uts.md) | Program mini dan laporan verifikasi |

## Cara menggunakan modul dalam 2 SKS

Setiap modul dirancang untuk **satu pertemuan 2 SKS atau 100 menit**. Jalur inti tiap modul mengikuti pola berikut:

| Tahap | Durasi umum | Tujuan |
|---|---:|---|
| Orientasi | 5–10 menit | menghubungkan materi dengan kasus teknik sipil |
| Konsep inti | 15–20 menit | mengenalkan satu gagasan utama tanpa terlalu banyak sintaks |
| Demonstrasi dosen | 15–20 menit | memperlihatkan alur dari input hingga hasil |
| Praktik terbimbing | 30–35 menit | mahasiswa meniru, mengubah, dan menjalankan kode |
| Uji dan refleksi | 10–15 menit | memeriksa hasil, menjelaskan program, dan membuat *exit ticket* |

Setiap modul memiliki tabel **Alur 100 menit** dan **Keluaran minimum**. Bagian berlabel **Pengayaan/PR** tidak harus diselesaikan saat tatap muka. Dengan pembagian ini, materi tetap lengkap sebagai bahan baca, tetapi kegiatan kelas tetap sederhana dan realistis.

Mahasiswa dapat bekerja berpasangan saat praktik, tetapi setiap mahasiswa tetap menjalankan dan menjelaskan programnya sendiri. Untuk menghemat waktu, dosen dapat membagikan workbook dengan judul kolom dan data awal; mahasiswa tetap membuat atau melengkapi algoritma dan kode. Contoh VBA ditempel pada **Visual Basic Editor** melalui `Alt+F11` → `Insert` → `Module`.

## Perangkat yang dibutuhkan

- Microsoft Excel desktop dengan dukungan macro/VBA;
- berkas kerja disimpan sebagai `.xlsm`;
- opsi minor: browser dan Google Colab, atau Python 3 dengan Jupyter Notebook; dan
- Git/GitHub opsional untuk menyimpan riwayat perubahan berkas Markdown dan kode.

> **Catatan keamanan:** aktifkan macro hanya untuk berkas yang sumbernya dipercaya. Jangan mengaktifkan macro pada lampiran yang tidak dikenal.

## Konvensi kode

Semua contoh VBA menggunakan `Option Explicit`. Nama variabel memakai istilah yang menjelaskan makna dan satuannya, misalnya `panjang_m` atau `debit_m3s`. Nilai input tidak ditanam langsung di banyak tempat; letakkan pada sel input atau konstanta yang jelas. Setiap hasil harus disertai satuan dan pemeriksaan kewajaran.

## Skema asesmen yang disarankan

| Komponen | Bobot |
|---|---:|
| Ketepatan algoritma dan hasil | 35% |
| Struktur, keterbacaan, dan dokumentasi kode | 20% |
| Validasi input dan penanganan galat | 15% |
| Verifikasi terhadap hitungan manual/kasus acuan | 20% |
| Kemampuan mendemonstrasikan dan menjelaskan | 10% |

## Etika penggunaan AI

AI boleh digunakan untuk menjelaskan pesan galat atau memberi alternatif kode jika dosen mengizinkan. Mahasiswa tetap wajib mampu menjelaskan setiap baris penting, menyebutkan bantuan yang dipakai, menguji program dengan datanya sendiri, dan bertanggung jawab atas hasil akhir.

---

Mulai dari [Modul 1](01-pengantar-komputasi.md).
