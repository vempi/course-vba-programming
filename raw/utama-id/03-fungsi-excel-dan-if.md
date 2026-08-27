# Modul 3 — Fungsi Excel dan Keputusan IF

## Capaian pembelajaran

Mahasiswa mampu:

- memakai fungsi agregasi dan logika yang tersedia di Excel;
- menulis `IF` tunggal dan majemuk;
- menggabungkan `IF` dengan `AND`/`OR`;
- mengklasifikasikan nilai mahasiswa menjadi A–E; dan
- menguji formula tepat pada nilai batas.

## Alur 100 menit

| Menit | Kegiatan |
|---:|---|
| 0–10 | Mengamati tabel nilai dan merumuskan keputusan |
| 10–25 | Fungsi dasar: `SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT` |
| 25–40 | `IF` tunggal dan operator perbandingan |
| 40–55 | `IF` majemuk dan urutan kondisi |
| 55–75 | Praktik penilaian A–E |
| 75–84 | Pengujian nilai batas dan input salah |
| 84–97 | Latihan singkat tiga soal |
| 97–100 | Checklist |

**Keluaran minimum:** satu tabel nilai mahasiswa yang menghitung rata-rata, status lulus, dan huruf A–E secara benar pada seluruh batas.

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

Perhatikan `>=`. Jika hanya memakai `>`, nilai tepat `60` akan salah diklasifikasikan.

## 3. Menggabungkan kondisi

Nilai dinyatakan valid jika berada di antara 0 dan 100:

```excel
=IF(AND(B2>=0,B2<=100),"Valid","Tidak valid")
```

Mahasiswa mendapat remedial jika nilai di bawah 60 **atau** kehadiran di bawah 75%:

```excel
=IF(OR(B2<60,C2<75%),"Remedial","Tidak remedial")
```

## 4. IF majemuk untuk A–E

Gunakan aturan latihan:

| Rentang | Huruf |
|---|---|
| 80–100 | A |
| 70–<80 | B |
| 60–<70 | C |
| 50–<60 | D |
| 0–<50 | E |

Validasi rentang dilakukan terlebih dahulu:

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

Uji minimal nilai akhir: `-1, 0, 49, 50, 59, 60, 69, 70, 79, 80, 100, 101`.

> Aturan nilai pada modul ini adalah data latihan, bukan kebijakan akademik resmi.

## 6. Kesalahan umum

- batas tidak mencakup nilai tepat, misalnya `>80`;
- urutan kondisi dimulai dari batas terendah sehingga semua nilai tinggi langsung masuk kategori pertama;
- angka disimpan sebagai teks;
- input di luar 0–100 tetap diberi huruf; dan
- formula terlalu panjang tanpa tabel aturan atau dokumentasi.

Untuk aturan yang sering berubah, tabel referensi dan fungsi pencarian biasanya lebih mudah dirawat daripada IF bertingkat. Itu dapat dipelajari sebagai pengayaan.

## Latihan singkat — 3 soal

### 1. Prediksi

Apa hasil formula berikut untuk `B2=80` dan mengapa?

```excel
=IF(B2>80,"A",IF(B2>70,"B","C"))
```

### 2. Praktik perbaikan

Perbaiki formula pada soal 1 agar batas A adalah `>=80`, batas B adalah `>=70`, dan input di luar 0–100 menghasilkan `Tidak valid`.

### 3. Jelaskan

Seorang mahasiswa mendapat tugas 100, UTS 50, dan UAS 50. Prediksi nilai akhir dan hurufnya. Jelaskan mengapa menguji hanya nilai 75 belum cukup untuk membuktikan formula A–E benar.

<details>
<summary>Kunci dan indikator pemahaman</summary>

1. Hasilnya `B`, karena `80>80` salah dan `80>70` benar. Ini menunjukkan dampak operator batas.
2. `=IF(OR(B2<0,B2>100),"Tidak valid",IF(B2>=80,"A",IF(B2>=70,"B","C")))`.
3. Nilai akhir `30+15+20=65`, sehingga C. Nilai 75 hanya menguji bagian tengah; nilai tepat di setiap batas, di luar rentang, dan ujung 0/100 juga harus diuji.

</details>

## Checklist

- [ ] Saya memahami argumen dan hasil fungsi.
- [ ] Saya dapat menulis IF tunggal dan majemuk.
- [ ] Saya sengaja memilih `>` atau `>=`.
- [ ] Saya menguji tepat pada setiap nilai batas.

## Bacaan lanjut

1. Michael Alexander & Dick Kusleika, *Microsoft Excel 365 Bible*, 2nd ed., Wiley, 2025.
2. Wayne L. Winston, *Microsoft Excel Data Analysis and Business Modeling*, 6th ed., Microsoft Press, 2019.
3. Bernard Liengme & Keith Hekman, *Liengme’s Guide to Excel 2016 for Scientists and Engineers*, Academic Press, 2019.
4. E. Joseph Billo, *Excel for Scientists and Engineers: Numerical Methods*, Wiley, 2007.

[← Modul 2](02-dasar-excel.md) · [Daftar modul](README.md) · [Modul 4 →](04-algoritma.md)
