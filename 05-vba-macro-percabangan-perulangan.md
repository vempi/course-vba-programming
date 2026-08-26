# Modul 5 — VBA: Macro Linier, Percabangan, dan Perulangan

## Capaian pembelajaran

Mahasiswa mampu:

- mengenali komponen Excel yang diakses VBA;
- membuka Visual Basic Editor dan memahami struktur `Sub`;
- mendeklarasikan variabel dengan tipe data yang sesuai;
- menjelaskan compile, runtime, statement, dan tiga jenis galat;
- merekam operasi spreadsheet menjadi macro serta membersihkan hasilnya;
- membuat program linier tanpa recorder;
- menggunakan `If...Then...ElseIf` dan `Select Case` untuk keputusan; serta
- menggunakan `For...Next` untuk memproses banyak baris dengan validasi.

## Alur 100 menit

| Menit | Kegiatan | Bagian |
|---:|---|---|
| 0–8 | Menelusuri objek Workbook–Worksheet–Range | §1 |
| 8–16 | Visual Basic Editor, module, `Sub`, dan statement | §2 |
| 16–26 | Variabel, tipe data, dan operator | §3 |
| 26–34 | Compile, runtime, dan tiga jenis galat | §4 |
| 34–44 | Merekam macro dan membersihkan hasilnya | §5 |
| 44–56 | Macro linier volume beton | §6 |
| 56–68 | Percabangan `If...ElseIf` dan validasi | §7–8 |
| 68–82 | Perulangan `For...Next` pada tabel data | §9 |
| 82–90 | Praktik klasifikasi 10 baris | §12 |
| 90–97 | Kuis tiga soal | Kuis |
| 97–100 | Simpan `.xlsm` dan checklist | Checklist |

**Keluaran minimum:** satu macro rekaman yang telah dibersihkan, satu macro linier buatan sendiri, dan satu macro berulang yang memproses sedikitnya sepuluh baris serta menangani nilai nonnumerik.

`Do While` (§10), pencarian baris terakhir (§11), dan ringkasan otomatis cukup dibaca atau dijadikan pengayaan.

## 1. Komponen Excel dari sudut pandang VBA

```text
Application
└── Workbook
    └── Worksheet
        └── Range/Cells
```

Contoh referensi lengkap:

```vb
ThisWorkbook.Worksheets("Volume").Range("B2").Value
```

- `ThisWorkbook` adalah workbook tempat kode disimpan.
- `Worksheets("Volume")` memilih lembar bernama Volume.
- `Range("B2")` memilih sel.
- `.Value` membaca atau mengubah nilainya.

Referensi lengkap lebih andal daripada mengandalkan workbook atau worksheet yang kebetulan aktif.

## 2. Menyiapkan VBA

1. Simpan workbook sebagai `.xlsm`.
2. Aktifkan tab Developer.
3. Tekan `Alt+F11`.
4. Pilih `Insert` → `Module`.
5. Tulis `Option Explicit` pada baris pertama.

Struktur prosedur:

```vb
Option Explicit

Sub NamaProsedur()
    ' statement dijalankan dari atas ke bawah
End Sub
```

Satu statement adalah satu instruksi. Tanda apostrof memulai komentar.

## 3. Variabel, tipe data, dan operator

| Tipe | Contoh | Kegunaan |
|---|---|---|
| `Double` | `2.75` | ukuran, debit, luas, hasil desimal |
| `Long` | `1500` | jumlah baris atau penghitung loop |
| `String` | `"Saluran A"` | nama, kode, keterangan |
| `Boolean` | `True`/`False` | status valid/tidak |
| `Date` | `#8/26/2026#` | tanggal pengukuran |
| `Variant` | bermacam nilai | fleksibel, gunakan hanya bila perlu |

```vb
Dim lebar_m As Double
Dim jumlahData As Long
Dim namaSaluran As String
Dim inputValid As Boolean
```

Gunakan `Double` untuk sebagian besar besaran teknik. Hindari `Integer` untuk nomor baris karena kapasitasnya terbatas; gunakan `Long`.

| Kelompok | Operator | Contoh |
|---|---|---|
| Aritmetika | `+ - * / ^` | `luas = lebar * tinggi` |
| Pembagian bulat | `\` | `7 \ 2` menghasilkan `3` |
| Sisa pembagian | `Mod` | `7 Mod 2` menghasilkan `1` |
| Perbandingan | `= <> < <= > >=` | `kedalaman > 0` |
| Logika | `And Or Not` | `lebar > 0 And tinggi > 0` |
| Gabung teks | `&` | `"Q = " & debit` |

Nama variabel menyertakan satuan — `panjang_m`, `debit_m3s` — sehingga kesalahan satuan terlihat saat membaca kode, bukan setelah hasilnya salah.

## 4. Compile, runtime, dan tiga jenis galat

- **Compile/check:** VBA memeriksa sintaks dan deklarasi sebelum prosedur dijalankan. Gunakan `Debug` → `Compile VBAProject`.
- **Runtime:** statement dijalankan dan berinteraksi dengan nilai atau objek aktual.

| Jenis galat | Kapan muncul | Contoh |
|---|---|---|
| Sintaks/compile | sebelum berjalan | `End Sub` hilang |
| Runtime | saat berjalan | worksheet yang dirujuk tidak ada |
| Logika | tidak pernah muncul | penjumlahan dipakai menggantikan perkalian |

Galat logika paling berbahaya karena program tetap menghasilkan angka. Satu-satunya cara menemukannya adalah membandingkan hasil dengan kasus acuan.

## 5. Macro Recorder

Rekam alur berikut:

1. mulai `Record Macro` dengan nama `FormatJudul`;
2. tulis judul `Panjang (m)`, `Lebar (m)`, dan `Volume (m³)` di A1:C1;
3. buat font tebal dan latar abu-abu;
4. hentikan rekaman; dan
5. buka kode hasil rekaman.

Hasil recorder mungkin berisi banyak `Select` dan `Selection`:

```vb
Range("A1:C1").Select
Selection.Font.Bold = True
Selection.Interior.Color = RGB(217, 217, 217)
```

Versi yang lebih langsung:

```vb
With ThisWorkbook.Worksheets("Volume").Range("A1:C1")
    .Font.Bold = True
    .Interior.Color = RGB(217, 217, 217)
End With
```

Recorder berguna untuk menemukan nama objek, property, dan method. Hasilnya adalah titik awal yang perlu dibaca dan dirapikan, bukan kode akhir.

## 6. Program linier

Program linier menjalankan langkah menerus tanpa keputusan atau loop.

Siapkan lembar `Volume`:

| Sel | Isi |
|---|---|
| A2/B2 | Panjang (m) / 10 |
| A3/B3 | Lebar (m) / 2 |
| A4/B4 | Tinggi (m) / 0,3 |
| A6 | Volume (m³) |

```vb
Option Explicit

Sub HitungVolumeLinier()
    Dim ws As Worksheet
    Dim panjang_m As Double
    Dim lebar_m As Double
    Dim tinggi_m As Double
    Dim volume_m3 As Double

    Set ws = ThisWorkbook.Worksheets("Volume")

    panjang_m = ws.Range("B2").Value
    lebar_m = ws.Range("B3").Value
    tinggi_m = ws.Range("B4").Value

    volume_m3 = panjang_m * lebar_m * tinggi_m

    ws.Range("B6").Value = volume_m3
    ws.Range("B6").NumberFormat = "0.000"
End Sub
```

Hasil acuan adalah `6,000 m³`.

Tambahkan faktor kehilangan 5% sebagai konstanta bernama, bukan angka yang ditanam di tengah rumus:

```vb
Const FAKTOR_KEHILANGAN As Double = 1.05
Dim volumePesan_m3 As Double

volumePesan_m3 = volume_m3 * FAKTOR_KEHILANGAN
ws.Range("B7").Value = volumePesan_m3
```

Uji program setelah setiap perubahan kecil. Jangan menunggu sampai banyak perubahan menumpuk.

## 7. Percabangan: program memilih jalur

Program linier tidak dapat menolak input buruk. Percabangan menambahkan kemampuan itu.

```vb
If lebar_m > 0 And kedalaman_m > 0 And kecepatan_ms >= 0 Then
    debit_m3s = lebar_m * kedalaman_m * kecepatan_ms
Else
    MsgBox "Dimensi harus > 0 dan kecepatan tidak boleh negatif.", vbExclamation
End If
```

Untuk beberapa kategori berjenjang, gunakan `ElseIf`. Kategori berikut adalah kategori latihan yang sama seperti pada Modul 3, kini dalam bentuk VBA:

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

Urutan kondisi penting, persis seperti pada `IF` bertingkat di Excel. Setelah program mengetahui `v` tidak kurang dari `0,3`, kondisi berikutnya cukup memeriksa `v <= 2`.

> Batas di atas hanya untuk latihan algoritma, bukan kriteria desain universal.

## 8. `Select Case` untuk pilihan diskret

`Select Case` cocok untuk kode jenis material atau kelas pilihan — lebih mudah dibaca daripada `ElseIf` berantai ketika yang dibandingkan adalah nilai diskret.

```vb
Function FaktorKehilangan(ByVal kodeMaterial As String) As Double
    Select Case UCase(Trim(kodeMaterial))
        Case "BETON"
            FaktorKehilangan = 1.05
        Case "PASIR"
            FaktorKehilangan = 1.1
        Case "BATU"
            FaktorKehilangan = 1.08
        Case Else
            FaktorKehilangan = 0
    End Select
End Function
```

`Trim` membuang spasi di awal/akhir dan `UCase` menyamakan huruf menjadi kapital, sehingga `" beton"` dan `"Beton"` diperlakukan sama. `Case Else` mengembalikan `0` sebagai penanda kode tidak dikenal — pemanggil wajib memeriksanya.

## 9. Perulangan `For...Next`

Gunakan loop ketika jumlah pengulangan sudah diketahui. Contoh berikut menghitung debit untuk data pada baris 2 sampai 11 di lembar `DataDebit`.

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

Dua hal yang perlu diperhatikan:

1. **`IsNumeric` diperiksa sebelum `CDbl`.** Tanpa urutan ini, input teks menimbulkan galat runtime *type mismatch* dan program berhenti di tengah tabel.
2. **Baris yang bermasalah ditandai, bukan menghentikan loop.** Ini adalah wujud kode dari cabang flowchart pada Modul 4 §6 yang kembali ke alur utama.

## 10. Perulangan `Do While` (pengayaan)

Gunakan `Do While` jika jumlah pengulangan bergantung pada kondisi, bukan pada jumlah data.

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

Setiap `Do While` harus bergerak menuju kondisi berhenti. Jika `nilaiKini` tidak berubah, loop berjalan tanpa akhir — gagasan kondisi berhenti yang sama seperti pada iterasi akar dua di Modul 2.

## 11. Menemukan baris terakhir otomatis (pengayaan)

Alih-alih menetapkan baris 11, VBA dapat mencari baris data terakhir:

```vb
Dim barisTerakhir As Long
barisTerakhir = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

For baris = 2 To barisTerakhir
    ' proses setiap baris
Next baris
```

Pastikan kolom acuan — di sini kolom A — terisi untuk setiap baris data.

## 12. Praktik terbimbing — klasifikasi rasio tulangan

Untuk latihan pemrograman, gunakan kategori berikut berdasarkan rasio `rho` dalam persen:

- `rho <= 0`: tidak valid;
- `0 < rho < 0,5`: rendah;
- `0,5 <= rho <= 2,5`: sedang; dan
- `rho > 2,5`: tinggi.

Buat tabel 10 data dan macro yang:

1. memeriksa apakah input numerik;
2. mengklasifikasikan setiap nilai dengan `If...ElseIf`;
3. menulis kategori pada kolom sebelahnya; dan
4. mewarnai sel input tidak valid.

```vb
ws.Cells(baris, "B").Interior.Color = RGB(255, 199, 206)
```

### Pengujian batas

| Jenis uji | Nilai `rho` (%) | Hasil yang diharapkan |
|---|---:|---|
| Tidak valid | -0,1 | tidak valid |
| Batas bawah | 0 | tidak valid |
| Tepat sebelum batas | 0,49 | rendah |
| Tepat pada batas | 0,50 | sedang |
| Tepat batas atas | 2,50 | sedang |
| Lewat batas | 2,51 | tinggi |
| Salah tipe | `abc` | bukan angka |

Ini tabel uji yang sama polanya dengan Modul 3 §7 — kali ini dijalankan oleh macro, bukan formula.

> Kategori ini hanya data latihan. Jangan menggunakannya untuk desain struktur.

### Pengayaan/PR

Ubah `HitungSemuaDebit` agar berhenti di baris terakhir secara otomatis, menghitung jumlah data valid dan tidak valid, lalu menampilkan ringkasan melalui `MsgBox` setelah loop selesai.

## Kuis berdampak — 3 soal

### 1. Prediksi

Macro recorder menghasilkan `Range("B2").Select` lalu `Selection.Value = 10`. Apa risiko jika worksheet lain sedang aktif? Tulis versi yang tidak bergantung pada seleksi.

### 2. Praktik perbaikan

Pada `HitungSemuaDebit`, seorang mahasiswa memindahkan `CDbl` ke atas sehingga dijalankan sebelum `IsNumeric`. Prediksi apa yang terjadi ketika baris ke-5 berisi teks, dan jelaskan mengapa akibatnya lebih buruk daripada sekadar satu baris salah.

### 3. Jelaskan

Kelompokkan tiga kejadian: `End Sub` hilang; worksheet `Volume` tidak ada; rumus memakai penjumlahan alih-alih perkalian. Mana galat sintaks/compile, runtime, dan logika? Mana yang paling berbahaya dan mengapa?

<details>
<summary>Kunci dan indikator pemahaman</summary>

1. Sel B2 pada sheet aktif dapat berubah — data di lembar lain tertimpa tanpa peringatan. Gunakan `ThisWorkbook.Worksheets("Volume").Range("B2").Value = 10`.
2. `CDbl("abc")` menimbulkan galat runtime *type mismatch*, program berhenti di baris ke-5, dan baris 6–11 **tidak pernah diproses**. Akibatnya bukan satu baris salah, melainkan sebagian tabel berisi hasil lama yang tampak sah.
3. `End Sub` hilang: sintaks/compile; sheet tidak ada: runtime; penjumlahan menggantikan perkalian: logika. Yang paling berbahaya adalah galat logika, karena dua jenis lainnya memberi pesan sedangkan galat logika menghasilkan angka yang tampak wajar.

</details>

## Checklist

- [ ] Workbook disimpan sebagai `.xlsm`.
- [ ] Saya dapat menelusuri Workbook–Worksheet–Range.
- [ ] Semua variabel dideklarasikan dengan tipe yang sesuai.
- [ ] Saya memahami setiap baris penting hasil recorder.
- [ ] Macro linier cocok dengan hitungan manual.
- [ ] Setiap percabangan menangani nilai batas secara sengaja.
- [ ] Teks diperiksa dengan `IsNumeric` sebelum dikonversi.
- [ ] Loop memiliki awal, akhir/kondisi berhenti, dan perubahan nilai.
- [ ] Input tidak valid ditandai tanpa menghentikan pemrosesan baris lain.

## Ringkasan

VBA menggerakkan Excel melalui hierarki objek Workbook–Worksheet–Range. Macro linier cukup untuk satu hitungan, percabangan menambahkan keputusan dan penolakan input buruk, dan perulangan menerapkan langkah yang sama pada banyak data. Ketiga struktur pada Modul 4 kini punya bentuk sintaksnya — dan urutan `IsNumeric` sebelum `CDbl` adalah pembeda antara program yang gagal dengan jelas dan program yang berhenti diam-diam di tengah tabel.

## Bacaan lanjut

1. Michael Alexander & Dick Kusleika, *Excel 2019 Power Programming with VBA*, Wiley, 2019.
2. Bill Jelen & Tracy Syrstad, *Microsoft Excel VBA and Macros (Office 2021 and Microsoft 365)*, Microsoft Press, 2022.
3. Steven Roman, *Writing Excel Macros with VBA*, 2nd ed., O'Reilly Media, 2002.
4. Bernard Liengme & Keith Hekman, *Liengme's Guide to Excel 2016 for Scientists and Engineers*, Academic Press, 2019.

[← Modul 4](04-algoritma-dan-verifikasi.md) · [Daftar modul](README.md) · [Modul 6 →](06-modularitas-data-dan-pengujian.md)
