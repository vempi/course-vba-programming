# Pengantar Komputer, Excel, dan VBA untuk Teknik Sipil

**Bahasa:** Indonesia · English version (segera hadir)

Paket materi Markdown untuk Program Sarjana Teknik Sipil, Departemen Teknik Sipil dan Lingkungan, Universitas Gadjah Mada. Materi menggunakan Excel dan VBA sebagai lingkungan utama, dengan demonstrasi opsional VBA AutoCAD.

## Desain pembelajaran

Setiap modul dirancang untuk **2 SKS atau 100 menit**. Alur pembelajaran dibuat sederhana:

```text
lihat kasus → prediksi → pelajari konsep → demonstrasi → praktik
            → uji hasil → jelaskan dengan kata-kata sendiri
```

Setiap pertemuan menghasilkan satu artefak kecil yang dapat didemonstrasikan dan ditutup dengan tiga soal berdampak tinggi:

1. **Prediksi:** mahasiswa menebak hasil sebelum menjalankan program.
2. **Praktik/perbaikan:** mahasiswa membuat atau memperbaiki artefak.
3. **Jelaskan:** mahasiswa menerangkan alasan, bukan menghafal sintaks.

## Daftar modul

| Modul | Materi | Produk minimum |
|---:|---|---|
| 1 | [Pengantar komputer, OS, dan pemrograman teknik sipil](01-pengantar-komputer-os.md) | peta kerja komputer dan dua demo script |
| 2 | [Dasar Excel dan fungsi](02-dasar-excel.md) | lembar hitungan dengan acuan dan fungsi yang benar |
| 3 | [Fungsi Excel dan keputusan IF](03-fungsi-excel-dan-if.md) | penilaian A–E yang teruji pada nilai batas |
| 4 | [Algoritma dan elemen-elemennya](04-algoritma.md) | IPO, pseudocode, flowchart, dan hitungan manual |
| 5 | [Pengenalan VBA dan macro linier](05-vba-dan-macro-linier.md) | macro rekaman yang dimodifikasi menjadi program |
| 6 | [Input–output, variabel, array, Function, dan Sub](06-input-output-dan-modularitas.md) | pengolahan beberapa baris data secara modular |
| 7 | [UserForm dan Control](07-userform-dan-control.md) | form input sederhana berbasis event |

## Pola waktu 100 menit

| Tahap | Durasi | Aktivitas |
|---|---:|---|
| Orientasi | 5–10 menit | kasus, tujuan, dan prediksi awal |
| Konsep inti | 15–25 menit | penjelasan singkat dan contoh konkret |
| Demonstrasi | 15–20 menit | dosen memperlihatkan proses lengkap |
| Praktik | 30–40 menit | mahasiswa membuat dan mengubah contoh |
| Kuis dan refleksi | 10–15 menit | tiga soal, pemeriksaan, dan penjelasan |

Bagian **pengayaan** tidak wajib selesai selama tatap muka. Dosen dapat menyediakan workbook dengan judul tabel dan data awal agar waktu kelas digunakan untuk memahami algoritma dan kode.

## Perangkat

- Microsoft Excel desktop dengan VBA;
- workbook disimpan sebagai `.xlsm`;
- AutoCAD dengan lingkungan VBA tersedia hanya untuk demo opsional; dan
- browser untuk membaca materi Markdown melalui GitHub.

> Aktifkan macro hanya dari sumber yang dipercaya. Simpan salinan data sebelum menjalankan macro yang mengubah banyak sel atau objek gambar.

## Konvensi

- Semua module VBA dimulai dengan `Option Explicit`.
- Nama variabel menyertakan makna dan satuan, misalnya `panjang_m`.
- Contoh formula Excel memakai nama fungsi berbahasa Inggris. Pemisah argumen dapat berupa koma atau titik koma, mengikuti pengaturan Excel.
- **F4** mengganti bentuk acuan relatif/absolut saat mengedit formula. **F5** membuka dialog *Go To*.
- Hasil program selalu diperiksa dengan kasus sederhana atau hitungan manual.

## Skema asesmen singkat

| Aspek | Bobot |
|---|---:|
| Artefak berfungsi dan hasil benar | 35% |
| Algoritma, satuan, dan validasi | 25% |
| Tiga soal kuis | 20% |
| Kemampuan mendemonstrasikan dan menjelaskan | 20% |

---

Mulai dari [Modul 1](01-pengantar-komputer-os.md).
