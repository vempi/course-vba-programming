# Modul 2 — Dasar Excel, Fungsi, dan Disiplin Satuan

## Capaian pembelajaran

Mahasiswa mampu:

- mengenali workbook, worksheet, baris, kolom, sel, range, dan formula bar;
- membedakan angka, teks, tanggal, logika, kosong, dan galat;
- membuat formula sederhana;
- menggunakan acuan relatif, absolut, dan campuran;
- menggunakan fungsi Excel yang umum;
- menjaga konsistensi satuan dan mendokumentasikan konversinya; serta
- menjelaskan nilai nol, galat, dan iterasi sederhana.

## Alur 100 menit

| Menit | Kegiatan | Bagian |
|---:|---|---|
| 0–10 | Orientasi antarmuka, alamat sel, dan jenis data | §1–2 |
| 10–26 | Formula serta acuan relatif/absolut | §3–4 |
| 26–42 | Demo tabel volume dan biaya | §5 |
| 42–55 | Fungsi `SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, `ROUND` | §6 |
| 55–68 | Audit satuan dan konversi eksplisit | §7 |
| 68–78 | Nol, sel kosong, teks, dan nilai galat | §8 |
| 78–86 | Ilustrasi proses iterasi | §9 |
| 86–97 | Kuis tiga soal | Kuis |
| 97–100 | Pemeriksaan dan checklist | Checklist |

**Keluaran minimum:** satu tabel volume/biaya yang dapat disalin ke bawah tanpa salah acuan, satu kotak ringkasan fungsi, satu baris audit satuan, serta satu tabel iterasi lima langkah.

## 1. Komponen utama Excel

- **Workbook:** satu berkas Excel.
- **Worksheet:** satu lembar di dalam workbook.
- **Cell:** perpotongan baris dan kolom, misalnya `B3`.
- **Range:** kumpulan sel, misalnya `B3:D10`.
- **Formula bar:** menampilkan isi asli sel.
- **Name box:** menampilkan atau menerima alamat sel/range.

Sel menyimpan nilai atau formula. Format hanya mengubah tampilan, bukan nilai dasarnya. Angka `0,15` yang ditampilkan sebagai `15%` tetap memiliki nilai numerik `0,15`.

## 2. Jenis data

| Isi | Jenis | Catatan |
|---|---|---|
| `12.5` | angka | dapat dihitung |
| `Saluran A` | teks | label/kode |
| `26/08/2026` | tanggal | disimpan sebagai nomor seri tanggal |
| `TRUE` | logika | benar/salah |
| sel tanpa isi | kosong | berbeda dari nol |
| `#DIV/0!` | galat | operasi tidak sah, misalnya membagi nol |

Jangan menulis satuan di dalam nilai seperti `12 m`. Simpan `12` sebagai angka dan letakkan `(m)` pada judul kolom. Sel yang berisi `12 m` adalah **teks**, dan teks tidak dapat dihitung.

## 3. Formula sederhana

Semua formula dimulai dengan `=`.

```excel
=B2*C2
=(B2+C2)/2
=PI()*B2^2/4
```

Gunakan tanda kurung agar urutan operasi jelas.

## 4. Acuan sel

Misalkan kolom B berisi volume dan sel `F1` berisi harga satuan.

| Formula | Jenis acuan | Saat disalin satu baris ke bawah |
|---|---|---|
| `=B2*F1` | relatif | menjadi `=B3*F2` |
| `=B2*$F$1` | absolut untuk F1 | menjadi `=B3*$F$1` |
| `=$B2*F$1` | campuran | kolom B dan baris 1 dikunci |

Saat kursor berada pada alamat sel di dalam formula, tekan **F4** untuk berputar melalui `F1`, `$F$1`, `F$1`, dan `$F1`. F5 bukan shortcut acuan absolut; F5 membuka *Go To*.

## 5. Praktik — volume dan biaya beton

Buat tabel:

| A | B | C | D | E |
|---|---:|---:|---:|---:|
| Segmen | Panjang (m) | Lebar (m) | Tinggi (m) | Volume (m³) |
| S1 | 10 | 0,2 | 0,3 | formula |
| S2 | 12 | 0,2 | 0,3 | formula |

Letakkan harga beton per m³ pada `H2`. Di kolom F hitung biaya:

```excel
=E2*$H$2
```

Salin formula ke bawah. Periksa bahwa alamat volume berubah, sedangkan `H2` tetap.

## 6. Fungsi Excel yang tersedia

Fungsi menerima input atau **argumen** dan mengembalikan hasil. Buat kotak ringkasan untuk kolom volume `E2:E11`:

```excel
=SUM(E2:E11)
=AVERAGE(E2:E11)
=MIN(E2:E11)
=MAX(E2:E11)
=COUNT(E2:E11)
=ROUND(AVERAGE(E2:E11),3)
```

`COUNT` menghitung sel numerik, sedangkan `COUNTA` menghitung sel yang tidak kosong. Perbedaan ini penting ketika data mengandung teks atau label — dan menjadi cara cepat menemukan sel yang tanpa sengaja tersimpan sebagai teks.

## 7. Satuan sebagai bagian dari hitungan

Kesalahan satuan menghasilkan angka yang tampak wajar tetapi salah. Ini jenis kesalahan yang paling sulit ditemukan karena program tidak memberi pesan apa pun.

Konversi yang sering dipakai:

```text
250 cm  = 250 / 100  = 2,5 m
500 L/s = 500 / 1000 = 0,5 m³/s
25 MPa  = 25 N/mm²   = 25.000 kN/m²
1 tf    = 9,80665 kN
```

Praktik yang baik:

- tulis satuan pada **judul kolom**, bukan di dalam sel nilai;
- konversi semua input ke satu sistem satuan **sebelum** menghitung;
- letakkan faktor konversi pada sel tersendiri yang diberi label, bukan diketik langsung ke dalam formula; dan
- tampilkan satuan pada hasil, tetapi simpan nilainya tetap sebagai angka.

### Audit satuan

Sebelum mempercayai sebuah tabel, telusuri satu baris dari input hingga hasil dan tuliskan satuannya di setiap langkah:

| Langkah | Nilai | Satuan | Periksa |
|---|---:|---|---|
| Panjang | 1.200 | cm | perlu dikonversi |
| Panjang (konversi) | 12 | m | `=B2/100`, faktor ada di sel berlabel |
| Lebar | 0,2 | m | sudah benar |
| Tinggi | 0,3 | m | sudah benar |
| Volume | 0,72 | m³ | m × m × m = m³ ✓ |

Jika satuan hasil tidak keluar sebagaimana mestinya, rumusnya salah — berapa pun angkanya.

> Faktor konversi yang ditulis langsung di dalam formula (`=B2/100`) sulit ditemukan ketika suatu saat harus diubah. Letakkan `100` pada sel berlabel `cm per m`, lalu rujuk sel itu.

## 8. Nol, kosong, dan galat

- Nol adalah angka yang sah, tetapi dapat membuat pembagian tidak terdefinisi.
- Sel kosong menyatakan belum ada data; jangan selalu menganggapnya nol.
- Teks yang tampak seperti angka dapat mengganggu perhitungan.
- `#DIV/0!`, `#VALUE!`, `#NAME?`, `#REF!`, dan `#N/A` memberi petunjuk jenis masalah.

Jangan langsung menyembunyikan semua galat. Temukan penyebabnya. `IFERROR` baru digunakan jika perilaku pengganti memang sudah dirancang.

## 9. Ilustrasi iterasi

Iterasi mengulang aturan hingga hasil cukup stabil. Untuk mendekati `√2`, gunakan:

```text
x_baru = (x_lama + 2/x_lama) / 2
```

Mulai dengan `x₀ = 1`. Buat tabel:

| Iterasi | x lama | x baru | perubahan absolut |
|---:|---:|---:|---:|
| 1 | 1 | `=(B2+2/B2)/2` | `=ABS(C2-B2)` |

Nilai `x baru` menjadi `x lama` pada baris berikutnya. Hentikan ketika perubahan lebih kecil dari toleransi, misalnya `0,000001`.

Dua gagasan di sini akan dipakai lagi: **kondisi berhenti** menjadi syarat setiap perulangan pada Modul 5, dan **toleransi** menjadi dasar pengujian bilangan pecahan pada Modul 6.

> Ini contoh numerik yang dihitung baris per baris. Hindari mengaktifkan iterasi workbook (*iterative calculation*) tanpa memahami circular reference dan kondisi berhentinya.

## Kuis berdampak — 3 soal

### 1. Prediksi

Di `C2` terdapat formula `=A2*$B$1`. Jika disalin ke `C5`, menjadi formula apa? Apa yang terjadi jika tanda `$` dihilangkan?

### 2. Praktik perbaikan

Formula kecepatan `=B2/C2` menghasilkan `#DIV/0!`. Tulis pemeriksaan yang membedakan `C2` kosong, `C2=0`, dan `C2>0`. Jangan sekadar menutup semua kondisi dengan `IFERROR`.

### 3. Jelaskan

Sebuah tabel menghitung volume dari panjang dalam sentimeter dan lebar/tinggi dalam meter, tanpa konversi. Hasilnya `72` dan tampak wajar. Jelaskan cara audit satuan menemukan kesalahan ini, lalu hitung dua langkah pertama iterasi akar dua dari `x₀=1` dan jelaskan mengapa proses memerlukan toleransi.

<details>
<summary>Kunci dan indikator pemahaman</summary>

1. Menjadi `=A5*$B$1`. Tanpa `$`, acuan B1 ikut bergeser menjadi B4.
2. Contoh: `=IF(C2="","Data luas belum diisi",IF(C2=0,"Luas tidak boleh nol",B2/C2))`.
3. Audit satuan: `cm × m × m` menghasilkan `cm·m²`, bukan `m³` — satuan hasil tidak konsisten meskipun angkanya terlihat masuk akal; nilai benar adalah `0,72 m³`, seratus kali lebih kecil. Iterasi: `x₁=1,5`; `x₂≈1,4166667`. Toleransi menentukan kapan perubahan sudah cukup kecil; tanpa kondisi berhenti, iterasi dapat berjalan terus.

</details>

## Checklist

- [ ] Saya membedakan nilai, format, dan formula.
- [ ] Saya dapat memilih acuan relatif/absolut dengan sengaja.
- [ ] Saya membedakan nol, kosong, teks, dan galat.
- [ ] Satuan tertulis di judul kolom dan faktor konversi ada di sel berlabel.
- [ ] Saya dapat menelusuri satuan dari input sampai hasil.
- [ ] Saya dapat menjelaskan aturan serta kondisi berhenti iterasi.

## Ringkasan

Excel menyimpan nilai, bukan tampilan; acuan menentukan apa yang berubah saat formula disalin; dan fungsi memberi hasil dari sekumpulan argumen. Di atas semua itu, satuan adalah bagian dari algoritma — hitungan yang satuannya tidak konsisten tetap salah walaupun angkanya tampak masuk akal.

## Bacaan lanjut

1. Michael Alexander & Dick Kusleika, *Microsoft Excel 365 Bible*, 2nd ed., Wiley, 2025.
2. Bernard Liengme & Keith Hekman, *Liengme's Guide to Excel 2016 for Scientists and Engineers*, Academic Press, 2019.
3. Ronald W. Larsen, *Engineering with Excel*, 5th ed., Pearson, 2017.
4. Wayne L. Winston, *Microsoft Excel Data Analysis and Business Modeling*, 6th ed., Microsoft Press, 2019.

[← Modul 1](01-pengantar-komputer-dan-komputasi.md) · [Daftar modul](README.md) · [Modul 3 →](03-keputusan-dan-pengujian-batas.md)
