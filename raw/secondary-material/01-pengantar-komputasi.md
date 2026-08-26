# Modul 1 — Pengantar Komputasi dalam Teknik Sipil

## Capaian pembelajaran

Setelah pertemuan ini, mahasiswa mampu:

- menjelaskan peran data, algoritma, program, dan validasi dalam pekerjaan teknik sipil;
- membedakan rumus Excel, macro VBA, dan program Python;
- menguraikan persoalan menjadi input–proses–output; dan
- membuat serta menjalankan macro VBA sederhana.

## Alur 100 menit

| Menit | Kegiatan | Bagian modul |
|---:|---|---|
| 0–10 | Pemantik: contoh hitungan sipil yang berulang di Excel | §1 |
| 10–25 | Alur masalah–algoritma–kode–validasi | §1–2 |
| 25–35 | Menyiapkan Excel dan Visual Basic Editor | §3 |
| 35–50 | Demonstrasi macro pertama | §4 |
| 50–75 | Praktik terbimbing volume pelat | §5 |
| 75–90 | Praktik mandiri volume sloof + 5% | §7 |
| 90–97 | Bandingkan hasil dengan hitungan manual | §5 dan §7 |
| 97–100 | *Exit ticket* dan checklist | §8–9 |

**Keluaran minimum:** satu workbook `.xlsm`, satu macro yang berjalan, dan satu hasil yang cocok dengan hitungan manual. Bagian Python pada §6 cukup ditunjukkan dosen selama 2–3 menit; mahasiswa tidak wajib menjalankannya di kelas.

## 1. Dari masalah teknik ke program

Komputasi teknik bukan sekadar mengetik rumus. Alurnya adalah:

```text
masalah nyata → asumsi/model → data input → algoritma → kode
             → hasil → validasi → keputusan teknik
```

Contoh: kita ingin menghitung volume pelat beton berbentuk balok.

- **Masalah:** berapa volume beton yang dibutuhkan?
- **Asumsi:** pelat berbentuk balok sempurna; kehilangan material belum dihitung.
- **Input:** panjang, lebar, dan tebal dalam meter.
- **Algoritma:** `volume = panjang × lebar × tebal`.
- **Output:** volume dalam m³.
- **Validasi:** hitung satu kasus secara manual dan periksa satuannya.

Program yang menghasilkan angka belum tentu benar. Rumus bisa salah, satuan bisa tidak konsisten, atau input bisa tidak masuk akal. Karena itu, validasi adalah bagian dari pemrograman—bukan pekerjaan tambahan setelah program selesai.

## 2. Analogi Excel, VBA, dan Python

| Kebutuhan | Excel | VBA | Python |
|---|---|---|---|
| Satu hitungan langsung | `=B2*C2*D2` | Bisa, tetapi belum perlu | Bisa, tetapi belum perlu |
| Mengulang 1.000 baris | Salin rumus | Loop/macro | Loop atau operasi array |
| Memeriksa input | Data Validation/`IF` | `If...Then` dan pesan | `if` dan exception |
| Membuat tombol proses | Form Control | Sangat sesuai | Perlu antarmuka lain |
| Analisis data besar/lanjut | Terbatas | Cukup untuk skala kecil–menengah | Sangat sesuai |

Prinsipnya sama: input dibaca, proses dijalankan, hasil ditulis. Perbedaannya terutama pada sintaks dan lingkungan kerja.

## 3. Menyiapkan Excel untuk VBA

1. Buka Excel desktop dan buat workbook kosong.
2. Simpan sebagai **Excel Macro-Enabled Workbook (`.xlsm`)**.
3. Tampilkan tab **Developer** melalui `File` → `Options` → `Customize Ribbon` → centang `Developer`.
4. Tekan `Alt+F11`, pilih `Insert` → `Module`.
5. Pastikan baris pertama kode adalah `Option Explicit`.

`Option Explicit` memaksa setiap variabel dideklarasikan. Ini membantu menangkap salah ketik nama variabel sebelum hasil keliru tersebar.

## 4. Demonstrasi 1 — macro pertama

Tempel kode berikut ke module, letakkan kursor di dalam prosedur, lalu tekan `F5`.

```vb
Option Explicit

Sub SapaTeknikSipil()
    MsgBox "Halo! Kita mulai komputasi teknik sipil.", _
           vbInformation, "Algoritma dan Pemrograman"
End Sub
```

Anatomi singkat:

- `Sub ... End Sub` membatasi sebuah prosedur;
- `MsgBox` menampilkan output sederhana;
- tanda `_` melanjutkan satu perintah ke baris berikutnya; dan
- teks diapit tanda petik ganda.

## 5. Demonstrasi 2 — volume pelat beton

Buat lembar kerja berikut.

| Sel | Isi |
|---|---|
| A2 | Panjang (m) |
| A3 | Lebar (m) |
| A4 | Tebal (m) |
| A6 | Volume (m³) |
| B2 | 6 |
| B3 | 4 |
| B4 | 0.15 |

Kemudian jalankan macro:

```vb
Option Explicit

Sub HitungVolumePelat()
    Dim panjang_m As Double
    Dim lebar_m As Double
    Dim tebal_m As Double
    Dim volume_m3 As Double

    panjang_m = Range("B2").Value
    lebar_m = Range("B3").Value
    tebal_m = Range("B4").Value

    volume_m3 = panjang_m * lebar_m * tebal_m

    Range("B6").Value = volume_m3
    Range("B6").NumberFormat = "0.000"
End Sub
```

Hasil acuan untuk data di atas adalah **3,600 m³**.

### Jejak eksekusi

| Langkah | Variabel/aksi | Nilai |
|---:|---|---:|
| 1 | baca `panjang_m` | 6 |
| 2 | baca `lebar_m` | 4 |
| 3 | baca `tebal_m` | 0,15 |
| 4 | hitung `volume_m3` | 3,6 |
| 5 | tulis ke B6 | 3,6 |

## 6. Perbandingan minor dengan Python (pengayaan singkat)

Di Google Colab/Jupyter, logika yang sama dapat ditulis sebagai berikut.

```python
panjang_m = 6.0
lebar_m = 4.0
tebal_m = 0.15

volume_m3 = panjang_m * lebar_m * tebal_m
print(f"Volume = {volume_m3:.3f} m³")
```

Sintaks berbeda, tetapi alur input–proses–output tetap sama. Dalam mata kuliah ini, VBA menjadi bahasa utama karena mudah dihubungkan dengan tabel Excel.

## 7. Praktik mandiri

Ubah lembar kerja dan macro agar menghitung volume balok sloof dari panjang, lebar, dan tinggi. Tambahkan output kebutuhan beton dengan faktor kehilangan 5%:

```text
volume_pemesanan = volume_geometris × 1,05
```

Uji minimal tiga kasus:

| Kasus | Panjang (m) | Lebar (m) | Tinggi (m) | Volume + 5% (m³) |
|---|---:|---:|---:|---:|
| Normal | 12 | 0,20 | 0,30 | hitung |
| Kecil | 1 | 0,15 | 0,20 | hitung |
| Batas | 0 | 0,20 | 0,30 | pikirkan apakah valid |

## 8. Pertanyaan refleksi dan *exit ticket*

Diskusikan pertanyaan 1–3 jika waktu tersedia. Jawab pertanyaan 4 dalam satu kalimat sebagai *exit ticket*.

1. Bagian mana yang merupakan model, algoritma, dan kode?
2. Mengapa nilai nol pada panjang perlu diperiksa?
3. Bagaimana memastikan bahwa semua input memakai meter, bukan sentimeter?
4. Kapan rumus Excel sudah cukup dan kapan VBA lebih bermanfaat?

## 9. Checklist hasil belajar

- [ ] Workbook tersimpan sebagai `.xlsm`.
- [ ] Macro dapat dijalankan tanpa pesan galat.
- [ ] Hasil cocok dengan hitungan manual.
- [ ] Nama variabel menjelaskan makna dan satuan.
- [ ] Saya dapat menjelaskan alur input–proses–output tanpa membaca kode.

## Ringkasan

Algoritma adalah urutan langkah untuk menyelesaikan masalah; kode adalah terjemahannya ke bahasa yang dapat dijalankan komputer. Dalam pekerjaan teknik, hasil program harus selalu disertai asumsi, satuan, dan validasi.

[← Daftar modul](README.md) · [Modul 2 →](02-representasi-masalah.md)
