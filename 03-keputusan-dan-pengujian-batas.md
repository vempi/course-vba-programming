# Modul 3 — Keputusan dan Pengujian Batas

## Capaian pembelajaran

Mahasiswa mampu:

- memakai fungsi agregasi dan logika yang tersedia di Excel;
- menulis `IF` tunggal dan majemuk;
- menggabungkan `IF` dengan `AND`/`OR`;
- mengklasifikasikan nilai ke dalam beberapa kategori berjenjang;
- menyusun tabel uji yang mencakup nilai normal, batas, di luar domain, dan salah tipe; serta
- menjelaskan mengapa pengujian di tengah rentang tidak membuktikan apa pun.

## Alur 100 menit

| Menit | Kegiatan | Bagian |
|---:|---|---|
| 0–10 | Mengamati tabel nilai dan merumuskan keputusan | §1 |
| 10–22 | Fungsi dasar dan operator perbandingan | §1–2 |
| 22–36 | `IF` tunggal, `AND`/`OR`, dan urutan kondisi | §2–3 |
| 36–50 | `IF` majemuk dan `IFS` untuk klasifikasi A–E | §4 |
| 50–64 | Praktik penilaian A–E | §5 |
| 64–74 | Klasifikasi besaran teknik | §6 |
| 74–86 | Menyusun dan menjalankan tabel uji batas | §7 |
| 86–97 | Latihan singkat tiga soal | Latihan |
| 97–100 | Checklist | Checklist |

**Keluaran minimum:** satu tabel nilai mahasiswa yang menghitung rata-rata, status valid, dan huruf A–E secara benar pada seluruh batas, disertai satu tabel uji berisi minimal tujuh kasus.

## 1. Fungsi yang sering digunakan

```excel
=SUM(B2:B11)
=AVERAGE(B2:B11)
=MIN(B2:B11)
=MAX(B2:B11)
=COUNT(B2:B11)
=ROUND(B2,2)
```

Fungsi menerima **argumen** dan mengembalikan **hasil**. Range `B2:B11` adalah satu argumen berupa kumpulan sel.

## 2. IF tunggal

Struktur umum:

```excel
=IF(kondisi, nilai_jika_benar, nilai_jika_salah)
```

Contoh kelulusan dengan batas `60`:

```excel
=IF(B2>=60,"Lulus","Belum lulus")
```

Perhatikan `>=`. Jika hanya memakai `>`, nilai tepat `60` akan salah diklasifikasikan. Operator batas adalah sumber kesalahan paling sering pada modul ini.

| Operator | Arti | Nilai tepat di batas |
|---|---|---|
| `>` | lebih besar | **tidak** termasuk |
| `>=` | lebih besar atau sama dengan | termasuk |
| `<` | lebih kecil | **tidak** termasuk |
| `<=` | lebih kecil atau sama dengan | termasuk |

## 3. Menggabungkan kondisi

Nilai dinyatakan valid jika berada di antara 0 dan 100:

```excel
=IF(AND(B2>=0,B2<=100),"Valid","Tidak valid")
```

Mahasiswa mendapat remedial jika nilai di bawah 60 **atau** kehadiran di bawah 75%:

```excel
=IF(OR(B2<60,C2<75%),"Remedial","Tidak remedial")
```

Gunakan `AND` jika semua syarat harus benar, `OR` jika salah satu syarat cukup, dan `NOT` untuk membalik nilai logika.

## 4. IF majemuk untuk A–E

Gunakan aturan latihan:

| Rentang | Huruf |
|---|---|
| 80–100 | A |
| 70–<80 | B |
| 60–<70 | C |
| 50–<60 | D |
| 0–<50 | E |

Validasi rentang dilakukan **terlebih dahulu**:

```excel
=IF(OR(B2<0,B2>100),"Tidak valid",
 IF(B2>=80,"A",
 IF(B2>=70,"B",
 IF(B2>=60,"C",
 IF(B2>=50,"D","E")))))
```

Pada Excel yang mendukung `IFS`, bagian klasifikasi dapat dibuat lebih mudah dibaca:

```excel
=IF(OR(B2<0,B2>100),"Tidak valid",
 IFS(B2>=80,"A",B2>=70,"B",B2>=60,"C",B2>=50,"D",TRUE,"E"))
```

Kondisi diperiksa dari batas tertinggi menuju terendah. Begitu kondisi benar ditemukan, kondisi berikutnya tidak dipakai.

> **Urutan menentukan hasil.** Jika kondisi disusun dari batas terendah (`B2>=50` lebih dulu), maka nilai `95` akan berhenti di cabang pertama dan mendapat `D`. Formula tetap berjalan, tidak ada pesan galat, dan seluruh kolom salah tanpa gejala.

## 5. Praktik tabel nilai

Buat kolom:

| Kolom | Isi |
|---|---|
| A | Nama |
| B–D | Tugas, UTS, UAS |
| E | Nilai akhir |
| F | Status valid |
| G | Huruf |

Gunakan bobot latihan `30% tugas + 30% UTS + 40% UAS`:

```excel
=B2*30%+C2*30%+D2*40%
```

> Aturan nilai pada modul ini adalah data latihan, bukan kebijakan akademik resmi.

## 6. Klasifikasi besaran teknik

Pola yang sama dipakai untuk mengelompokkan besaran teknik. Contoh kategori kecepatan aliran:

| Kecepatan `v` (m/s) | Kategori |
|---|---|
| `v < 0` | input tidak valid |
| `0 ≤ v < 0,3` | rendah |
| `0,3 ≤ v ≤ 2,0` | rentang pengamatan |
| `v > 2,0` | tinggi |

```excel
=IF(B2<0,"Input tidak valid",
 IF(B2<0.3,"Rendah",
 IF(B2<=2,"Rentang pengamatan","Tinggi")))
```

Perhatikan bahwa setelah program mengetahui `v` tidak kurang dari `0,3`, kondisi berikutnya cukup memeriksa `v <= 2`. Kondisi tidak perlu mengulang batas bawah yang sudah tersaring di cabang sebelumnya.

> Batas di atas hanya untuk latihan algoritma, bukan kriteria desain universal. Penetapan batas desain merujuk standar dan kondisi material saluran yang berlaku.

## 7. Pengujian batas

Percabangan paling sering salah **tepat di nilai batas**, dan kesalahan itu tidak terlihat jika pengujian hanya memakai nilai di tengah rentang. Nilai `75` akan menghasilkan `B` baik pada formula yang benar maupun pada formula yang batasnya keliru.

Susun tabel uji sebelum menyatakan formula selesai:

| Jenis uji | Nilai | Hasil yang diharapkan |
|---|---:|---|
| Di luar domain (bawah) | -1 | Tidak valid |
| Ujung domain | 0 | E |
| Tepat sebelum batas | 49 | E |
| Tepat pada batas | 50 | D |
| Tepat sebelum batas | 79 | B |
| Tepat pada batas | 80 | A |
| Ujung domain | 100 | A |
| Di luar domain (atas) | 101 | Tidak valid |
| Salah tipe | `abc` | Tidak valid |

Aturannya sederhana: **untuk setiap batas, uji tepat di batas dan tepat satu langkah sebelumnya.** Tambahkan kedua ujung domain dan minimal satu input salah tipe.

Catat hasilnya sebagai bukti, bukan sekadar dilihat sekilas:

| ID | Input | Harapan | Aktual | Status |
|---|---:|---|---|---|
| T1 | 79 | B | | |
| T2 | 80 | A | | |
| T3 | 101 | Tidak valid | | |

Kolom **aktual** dan **status** diisi setelah formula dijalankan. Pola pencatatan ini berkembang menjadi pengujian otomatis pada Modul 6.

## 8. Kesalahan umum

- batas tidak mencakup nilai tepat, misalnya `>80` padahal seharusnya `>=80`;
- urutan kondisi dimulai dari batas terendah sehingga semua nilai tinggi masuk kategori pertama;
- angka disimpan sebagai teks sehingga perbandingan tidak bekerja seperti yang diduga;
- input di luar 0–100 tetap diberi huruf karena validasi diletakkan setelah klasifikasi; dan
- formula panjang tanpa tabel aturan atau dokumentasi.

Untuk aturan yang sering berubah, tabel referensi dan fungsi pencarian biasanya lebih mudah dirawat daripada `IF` bertingkat. Itu dapat dipelajari sebagai pengayaan.

## Latihan singkat — 3 soal

### 1. Prediksi

Apa hasil formula berikut untuk `B2=80` dan mengapa?

```excel
=IF(B2>80,"A",IF(B2>70,"B","C"))
```

### 2. Praktik perbaikan

Perbaiki formula pada soal 1 agar batas A adalah `>=80`, batas B adalah `>=70`, dan input di luar 0–100 menghasilkan `Tidak valid`. Lalu tuliskan tiga nilai uji yang membuktikan perbaikan itu bekerja.

### 3. Jelaskan

Seorang mahasiswa mendapat tugas 100, UTS 50, dan UAS 50. Prediksi nilai akhir dan hurufnya. Jelaskan mengapa menguji hanya nilai 75 belum cukup untuk membuktikan formula A–E benar, dan sebutkan nilai uji minimum yang Anda perlukan.

<details>
<summary>Kunci dan indikator pemahaman</summary>

1. Hasilnya `B`, karena `80>80` salah dan `80>70` benar. Ini menunjukkan dampak operator batas.
2. `=IF(OR(B2<0,B2>100),"Tidak valid",IF(B2>=80,"A",IF(B2>=70,"B","C")))`. Nilai uji yang membuktikan: `79` → B, `80` → A, `101` → Tidak valid.
3. Nilai akhir `30+15+20=65`, sehingga C. Nilai 75 hanya menguji bagian tengah satu kategori dan akan lulus walaupun batas ditulis salah. Minimum yang diperlukan: setiap batas (50, 60, 70, 80) beserta nilai tepat di bawahnya (49, 59, 69, 79), kedua ujung domain (0 dan 100), dan nilai di luar domain (-1, 101).

</details>

## Checklist

- [ ] Saya memahami argumen dan hasil fungsi.
- [ ] Saya dapat menulis IF tunggal dan majemuk.
- [ ] Saya sengaja memilih `>` atau `>=` dan dapat menjelaskan alasannya.
- [ ] Validasi domain diletakkan sebelum klasifikasi.
- [ ] Saya menguji tepat pada setiap nilai batas dan satu langkah sebelumnya.
- [ ] Hasil pengujian saya catat sebagai tabel, bukan hanya dilihat sekilas.

## Ringkasan

Percabangan menerjemahkan aturan teknik atau akademik menjadi keputusan yang dapat dijalankan. Yang menentukan benar-salahnya bukan panjang formula, melainkan pilihan operator batas, urutan kondisi, dan letak validasi. Karena kesalahan batas tidak memberi pesan galat, satu-satunya bukti bahwa formula benar adalah tabel uji yang menyentuh setiap batas.

## Bacaan lanjut

1. Michael Alexander & Dick Kusleika, *Microsoft Excel 365 Bible*, 2nd ed., Wiley, 2025.
2. Wayne L. Winston, *Microsoft Excel Data Analysis and Business Modeling*, 6th ed., Microsoft Press, 2019.
3. Bernard Liengme & Keith Hekman, *Liengme's Guide to Excel 2016 for Scientists and Engineers*, Academic Press, 2019.
4. E. Joseph Billo, *Excel for Scientists and Engineers: Numerical Methods*, Wiley, 2007.

[← Modul 2](02-excel-fungsi-dan-satuan.md) · [Daftar modul](README.md) · [Modul 4 →](04-algoritma-dan-verifikasi.md)
