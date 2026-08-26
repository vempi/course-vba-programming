# Modul 3 — Struktur Kontrol

## Capaian pembelajaran

Mahasiswa mampu:

- menggunakan `If...Then...Else` dan `Select Case` untuk keputusan;
- menggunakan `For...Next` dan `Do While` untuk perulangan;
- memvalidasi input sebelum perhitungan; dan
- menerapkan struktur kontrol pada kasus teknik sipil sederhana.

## Alur 100 menit

| Menit | Kegiatan | Bagian modul |
|---:|---|---|
| 0–10 | Prediksi hasil beberapa kondisi `If` | §1–2 |
| 10–25 | Demonstrasi percabangan dan nilai batas | §1–3 |
| 25–40 | Demonstrasi loop pada tiga baris data | §4 |
| 40–50 | Menambahkan `IsNumeric` dan validasi domain | §2 dan §4 |
| 50–80 | Praktik klasifikasi rasio pada 10 data | §7 |
| 80–93 | Menguji nilai normal, batas, dan salah tipe | §8 |
| 93–100 | Perbaikan, checklist, dan *exit ticket* | Checklist dan Ringkasan |

**Keluaran minimum:** macro yang memproses 10 baris, membedakan sedikitnya tiga kategori, serta menangani nilai nonnumerik dan nilai di luar domain. `Do While`, pencarian baris terakhir, dan ringkasan otomatis cukup dibaca atau dijadikan pengayaan.

## 1. Percabangan: program memilih jalur

Misalkan kecepatan aliran dinilai dengan kriteria pembelajaran berikut:

| Kecepatan `v` | Kategori demonstrasi |
|---|---|
| `v < 0` | input tidak valid |
| `0 ≤ v < 0,3` m/s | rendah |
| `0,3 ≤ v ≤ 2,0` m/s | rentang pengamatan |
| `v > 2,0` m/s | tinggi |

> Batas di atas hanya untuk latihan algoritma, bukan kriteria desain universal. Pemilihan batas desain harus merujuk standar dan kondisi material saluran yang berlaku.

```vb
Function KategoriKecepatan(ByVal kecepatan_ms As Double) As String
    If kecepatan_ms < 0 Then
        KategoriKecepatan = "Input tidak valid"
    ElseIf kecepatan_ms < 0.3 Then
        KategoriKecepatan = "Rendah"
    ElseIf kecepatan_ms <= 2# Then
        KategoriKecepatan = "Rentang pengamatan"
    Else
        KategoriKecepatan = "Tinggi"
    End If
End Function
```

Urutan kondisi penting. Setelah program mengetahui `v` tidak kurang dari `0,3`, kondisi berikutnya cukup memeriksa `v <= 2`.

## 2. Operator logika dan validasi

Sebuah penampang hanya dapat dihitung jika semua dimensinya positif:

```vb
If lebar_m > 0 And kedalaman_m > 0 And kecepatan_ms >= 0 Then
    debit_m3s = lebar_m * kedalaman_m * kecepatan_ms
Else
    MsgBox "Dimensi harus > 0 dan kecepatan tidak boleh negatif.", vbExclamation
End If
```

Gunakan `And` jika semua syarat harus benar, `Or` jika salah satu syarat cukup, dan `Not` untuk membalik nilai logika.

## 3. `Select Case` untuk pilihan diskret

`Select Case` cocok untuk kode jenis material atau kelas pilihan:

```vb
Function FaktorKehilangan(ByVal kodeMaterial As String) As Double
    Select Case UCase(Trim(kodeMaterial))
        Case "BETON"
            FaktorKehilangan = 1.05
        Case "PASIR"
            FaktorKehilangan = 1.10
        Case "BATU"
            FaktorKehilangan = 1.08
        Case Else
            FaktorKehilangan = 0
    End Select
End Function
```

`Trim` membuang spasi di awal/akhir dan `UCase` menyamakan huruf menjadi kapital.

## 4. Perulangan `For...Next`

Gunakan loop ketika jumlah pengulangan sudah diketahui. Contoh berikut menghitung debit untuk data pada baris 2 sampai 11.

| Kolom | Isi |
|---|---|
| A | Nama penampang |
| B | Luas (m²) |
| C | Kecepatan (m/s) |
| D | Debit (m³/s) |
| E | Status |

```vb
Option Explicit

Sub HitungSemuaDebit()
    Dim ws As Worksheet
    Dim baris As Long
    Dim luas_m2 As Double
    Dim kecepatan_ms As Double

    Set ws = ThisWorkbook.Worksheets("DataDebit")

    For baris = 2 To 11
        If IsNumeric(ws.Cells(baris, "B").Value) And _
           IsNumeric(ws.Cells(baris, "C").Value) Then

            luas_m2 = CDbl(ws.Cells(baris, "B").Value)
            kecepatan_ms = CDbl(ws.Cells(baris, "C").Value)

            If luas_m2 > 0 And kecepatan_ms >= 0 Then
                ws.Cells(baris, "D").Value = luas_m2 * kecepatan_ms
                ws.Cells(baris, "E").Value = "OK"
            Else
                ws.Cells(baris, "D").ClearContents
                ws.Cells(baris, "E").Value = "Nilai tidak valid"
            End If
        Else
            ws.Cells(baris, "D").ClearContents
            ws.Cells(baris, "E").Value = "Bukan angka"
        End If
    Next baris
End Sub
```

Perhatikan urutannya: `IsNumeric` diperiksa sebelum `CDbl`. Tanpa urutan ini, input teks dapat menimbulkan galat tipe data.

## 5. Perulangan `Do While` (pengayaan konsep)

Gunakan `Do While` jika jumlah pengulangan bergantung pada kondisi. Contoh berikut mencari berapa tahun hingga suatu kuantitas berlipat dua pada pertumbuhan tetap. Ini adalah latihan struktur kontrol, bukan model prediksi teknik yang lengkap.

```vb
Sub SimulasiPertumbuhan()
    Dim nilaiAwal As Double
    Dim nilaiKini As Double
    Dim laju As Double
    Dim tahun As Long

    nilaiAwal = 100
    nilaiKini = nilaiAwal
    laju = 0.05
    tahun = 0

    Do While nilaiKini < 2 * nilaiAwal
        nilaiKini = nilaiKini * (1 + laju)
        tahun = tahun + 1
    Loop

    MsgBox "Melebihi dua kali nilai awal setelah " & tahun & " tahun."
End Sub
```

Setiap `Do While` harus bergerak menuju kondisi berhenti. Jika `tahun` atau `nilaiKini` tidak berubah, loop dapat berjalan tanpa akhir.

## 6. Menemukan baris terakhir otomatis (pengayaan)

Alih-alih menetapkan baris 11, VBA dapat mencari baris data terakhir:

```vb
Dim barisTerakhir As Long
barisTerakhir = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

For baris = 2 To barisTerakhir
    ' proses setiap baris
Next baris
```

Pastikan kolom acuan (di sini kolom A) terisi untuk setiap baris data.

## 7. Praktik terbimbing — klasifikasi rasio tulangan

Untuk latihan pemrograman, gunakan kategori buatan berikut berdasarkan rasio `rho` dalam persen:

- `rho <= 0`: tidak valid;
- `0 < rho < 0,5`: rendah;
- `0,5 <= rho <= 2,5`: sedang; dan
- `rho > 2,5`: tinggi.

Buat tabel 10 data dan macro yang:

1. memeriksa apakah input numerik;
2. mengklasifikasikan setiap nilai dengan `If...ElseIf`;
3. menulis kategori pada kolom sebelahnya; dan
4. mewarnai sel input tidak valid dengan warna merah muda.

Potongan pewarnaan:

```vb
ws.Cells(baris, "B").Interior.Color = RGB(255, 199, 206)
```

> Kategori ini hanya data latihan. Jangan menggunakannya untuk desain struktur.

## 8. Pengujian batas

Percabangan paling sering salah tepat di nilai batas. Uji paling sedikit:

| Jenis uji | Nilai `rho` (%) | Hasil yang diharapkan |
|---|---:|---|
| Tidak valid | -0,1 | tidak valid |
| Batas bawah | 0 | tidak valid |
| Tepat sebelum batas | 0,49 | rendah |
| Tepat pada batas | 0,50 | sedang |
| Tepat batas atas | 2,50 | sedang |
| Lewat batas | 2,51 | tinggi |
| Salah tipe | `abc` | bukan angka |

## 9. Tantangan pengayaan/PR

Ubah `HitungSemuaDebit` agar:

- berhenti di baris terakhir secara otomatis;
- menghitung jumlah data valid dan tidak valid; dan
- menampilkan ringkasan melalui `MsgBox` setelah loop selesai.

## Checklist

- [ ] Setiap percabangan menangani nilai batas secara sengaja.
- [ ] Teks diperiksa dengan `IsNumeric` sebelum dikonversi.
- [ ] Loop memiliki awal, akhir/kondisi berhenti, dan perubahan nilai.
- [ ] Input tidak valid tidak menghasilkan angka palsu.
- [ ] Saya menguji nilai normal, batas, dan salah.

## Ringkasan

Percabangan mengatur keputusan, sedangkan perulangan mengotomasi langkah yang berulang. Validasi harus dilakukan sebelum perhitungan sehingga program gagal dengan jelas dan aman, bukan diam-diam menghasilkan nilai yang menyesatkan.

[← Modul 2](02-representasi-masalah.md) · [Daftar modul](README.md) · [Modul 4 →](04-struktur-data.md)
