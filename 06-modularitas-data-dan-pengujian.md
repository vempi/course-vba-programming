# Modul 6 — Modularitas, Struktur Data, Debugging, dan Pengujian

## Capaian pembelajaran

Mahasiswa mampu:

- membaca dan menulis data melalui `Cells`, `Range`, `InputBox`, dan `MsgBox`;
- mendeklarasikan variabel dengan tipe dan ruang lingkup yang sesuai;
- memilih struktur data sesuai bentuk masalah dan menggunakan array;
- membedakan `Function` dan `Sub` serta memecah program menjadi bagian terfokus;
- menulis dokumentasi singkat tentang parameter, hasil, dan satuan;
- menelusuri kode dengan breakpoint, *step execution*, dan Immediate Window;
- menangani galat yang dapat diperkirakan; serta
- menyusun pengujian sederhana dengan hasil yang diketahui.

## Alur 100 menit

| Menit | Kegiatan | Bagian |
|---:|---|---|
| 0–8 | Mengidentifikasi input–proses–output macro sebelumnya | §1 |
| 8–18 | Variabel, ruang lingkup, `ByVal`/`ByRef` | §2 |
| 18–30 | Array satu dimensi dan blok Excel | §3–4 |
| 30–44 | `Function` dan `Sub`: memisahkan rumus dari alur Excel | §6–7 |
| 44–56 | Debugging: compile, breakpoint, `F8`, Immediate Window | §9 |
| 56–64 | Validasi dan penanganan galat terarah | §10 |
| 64–80 | Pengujian dengan nilai acuan | §11 |
| 80–90 | Praktik `VolumePipa` beserta ujinya | §12 |
| 90–97 | Kuis tiga soal | Kuis |
| 97–100 | Compile dan checklist | Checklist |

**Keluaran minimum:** satu `Function` terdokumentasi beserta satu `Sub` yang memakainya untuk memproses sedikitnya lima baris, dan tiga pengujian yang mencakup kasus normal, batas, serta tidak valid.

### Batas materi inti

Fokus praktik kelas adalah **array satu dimensi, pemisahan `Function`/`Sub`, dan pengujian**. Array dua dimensi, `Collection`, `Dictionary`, padanan Python, serta *error handler* lengkap tersedia agar materi utuh sebagai bahan baca, tetapi cukup didemonstrasikan dosen.

### Satu benang merah

Gunakan kasus pipa dari awal sampai akhir: buat fungsi luas, pakai fungsi itu untuk volume, sengaja masukkan kesalahan, telusuri dengan debugger, lalu buktikan perbaikannya dengan pengujian.

## 1. Input dan output

### `Range` dan `Cells`

```vb
Dim ws As Worksheet
Dim nilai As Double

Set ws = ThisWorkbook.Worksheets("DataVolume")

nilai = ws.Range("B2").Value
ws.Cells(2, "F").Value = nilai * 2
```

`Range("B2")` mudah dibaca untuk alamat tetap. `Cells(baris, kolom)` cocok ketika nomor baris atau kolom berubah di dalam loop.

### `InputBox` dan `MsgBox`

```vb
Sub DemoInputOutput()
    Dim namaProyek As String

    namaProyek = InputBox("Masukkan nama proyek:")

    If namaProyek = "" Then
        MsgBox "Nama belum diisi.", vbExclamation
    Else
        MsgBox "Proyek: " & namaProyek, vbInformation
    End If
End Sub
```

`InputBox` selalu mengembalikan teks. Periksa dengan `IsNumeric` sebelum mengubahnya menjadi angka dengan `CDbl` — aturan yang sama seperti pada Modul 5.

## 2. Ruang lingkup, `ByVal`, dan `ByRef`

Variabel local dideklarasikan di dalam prosedur dan hanya tersedia di sana:

```vb
Sub ContohLocal()
    Dim total_m3 As Double
    total_m3 = 10
End Sub
```

Variabel public dideklarasikan di bagian atas module:

```vb
Option Explicit
Public namaProyekAktif As String
```

Gunakan variabel local sebagai pilihan awal. Variabel public berguna ketika keadaan benar-benar perlu dibagi antarmodule atau form, tetapi nilainya lebih sulit ditelusuri karena dapat berubah dari banyak tempat.

Untuk parameter:

- `ByVal` memberikan **salinan** nilai; perubahan di dalam prosedur tidak mengubah variabel pemanggil.
- `ByRef` memberikan **referensi**; prosedur dapat mengubah variabel pemanggil.

Gunakan `ByVal` sebagai pilihan awal untuk parameter hitungan. Fungsi yang tidak mengubah apa pun di luar dirinya jauh lebih mudah diuji.

## 3. Memilih struktur data

Satu variabel menyimpan satu nilai. Data teknik biasanya berupa deret curah hujan, koordinat titik, daftar material, atau pasangan kode–nilai.

| Bentuk data | VBA | Python | Contoh teknik sipil |
|---|---|---|---|
| Deret berurutan, ukuran tetap | array | `list`/`tuple` | hujan 12 bulan |
| Tabel baris–kolom | array 2D atau `Range` | NumPy array | dimensi beberapa segmen |
| Daftar objek dinamis | `Collection` | `list` | nama stasiun |
| Pasangan kunci–nilai | `Dictionary` | `dict` | kode material–berat jenis |

## 4. Array satu dimensi

```vb
Option Explicit

Sub DemoArrayHujan()
    Dim hujan_mm(1 To 5) As Double
    Dim i As Long
    Dim total_mm As Double
    Dim rataRata_mm As Double

    hujan_mm(1) = 12.5
    hujan_mm(2) = 0
    hujan_mm(3) = 31.2
    hujan_mm(4) = 8.4
    hujan_mm(5) = 17.9

    For i = LBound(hujan_mm) To UBound(hujan_mm)
        total_mm = total_mm + hujan_mm(i)
    Next i

    rataRata_mm = total_mm / (UBound(hujan_mm) - LBound(hujan_mm) + 1)
    MsgBox "Rata-rata = " & Format(rataRata_mm, "0.00") & " mm"
End Sub
```

`LBound` dan `UBound` membuat loop mengikuti batas array, sehingga menambah data tidak memaksa mengubah angka di banyak tempat.

## 5. Array dinamis dan blok Excel (demonstrasi dosen)

Jika ukuran baru diketahui saat program berjalan, gunakan `ReDim`. Satu `Range.Value` juga dapat dipindahkan ke array 2D sekaligus — umumnya jauh lebih cepat daripada membaca sel satu per satu.

```vb
Sub HitungVolumeBanyakSegmen()
    Dim ws As Worksheet
    Dim data As Variant
    Dim hasil() As Variant
    Dim i As Long

    Set ws = ThisWorkbook.Worksheets("Segmen")
    data = ws.Range("B2:D11").Value  ' panjang, lebar, tinggi
    ReDim hasil(1 To UBound(data, 1), 1 To 1)

    For i = 1 To UBound(data, 1)
        If IsNumeric(data(i, 1)) And IsNumeric(data(i, 2)) And _
           IsNumeric(data(i, 3)) Then
            hasil(i, 1) = CDbl(data(i, 1)) * CDbl(data(i, 2)) * _
                          CDbl(data(i, 3))
        Else
            hasil(i, 1) = "Input salah"
        End If
    Next i

    ws.Range("E2:E11").Value = hasil
End Sub
```

Indeks pertama adalah baris array dan indeks kedua adalah kolom array.

### `Collection` dan `Dictionary` (pengayaan)

`Collection` menyimpan daftar yang ukurannya dapat bertambah. `Dictionary` menyimpan pasangan kunci–nilai dan cocok untuk pencarian berdasarkan kode:

```vb
Sub DemoDictionaryMaterial()
    Dim beratJenis As Object
    Dim kode As String

    Set beratJenis = CreateObject("Scripting.Dictionary")
    beratJenis.Add "BETON", 24#
    beratJenis.Add "BAJA", 78.5
    beratJenis.Add "AIR", 9.81

    kode = UCase(Trim(ThisWorkbook.Worksheets("Material").Range("B2").Value))

    If beratJenis.Exists(kode) Then
        ThisWorkbook.Worksheets("Material").Range("B3").Value = beratJenis(kode)
    Else
        ThisWorkbook.Worksheets("Material").Range("B3").ClearContents
    End If
End Sub
```

`Scripting.Dictionary` tersedia pada Excel Windows. Untuk kompatibilitas lintas platform, array tabel atau `Collection` menjadi alternatif.

## 6. Modularitas

Program yang baik dibagi menjadi bagian kecil dengan satu tanggung jawab.

```text
Sub utama
├── membaca dan memvalidasi input
├── memanggil fungsi hitungan
└── menulis hasil
```

`Sub` menjalankan tindakan, sedangkan `Function` mengembalikan nilai. Fungsi kecil mudah diuji tanpa harus menjalankan seluruh workbook.

## 7. Function dan Sub

`Function` menerima parameter dan mengembalikan satu nilai:

```vb
Option Explicit

' Menghitung luas lingkaran.
' diameter_m: diameter dalam meter, harus > 0.
' Hasil: luas dalam m²; mengembalikan -1 jika input tidak valid.
Public Function LuasLingkaran(ByVal diameter_m As Double) As Double
    Const PI As Double = 3.14159265358979

    If diameter_m <= 0 Then
        LuasLingkaran = -1
        Exit Function
    End If

    LuasLingkaran = PI * diameter_m ^ 2 / 4
End Function
```

Karena `Public Function` berada di module standar, fungsi ini juga dapat dipakai langsung pada sel Excel: `=LuasLingkaran(B2)`.

`Sub` mengatur alur: membaca tabel, memanggil fungsi, dan menulis hasil.

```vb
Public Function VolumeBalok( _
    ByVal panjang_m As Double, _
    ByVal lebar_m As Double, _
    ByVal tinggi_m As Double) As Double

    If panjang_m <= 0 Or lebar_m <= 0 Or tinggi_m <= 0 Then
        VolumeBalok = -1
    Else
        VolumeBalok = panjang_m * lebar_m * tinggi_m
    End If
End Function

Public Sub HitungTabelVolume()
    Dim ws As Worksheet
    Dim data As Variant
    Dim hasil() As Variant
    Dim i As Long
    Dim volume_m3 As Double

    Set ws = ThisWorkbook.Worksheets("DataVolume")
    data = ws.Range("B2:D6").Value
    ReDim hasil(1 To UBound(data, 1), 1 To 2)

    For i = 1 To UBound(data, 1)
        If IsNumeric(data(i, 1)) And IsNumeric(data(i, 2)) And _
           IsNumeric(data(i, 3)) Then

            volume_m3 = VolumeBalok(CDbl(data(i, 1)), _
                                    CDbl(data(i, 2)), _
                                    CDbl(data(i, 3)))

            If volume_m3 >= 0 Then
                hasil(i, 1) = volume_m3
                hasil(i, 2) = "OK"
            Else
                hasil(i, 1) = Empty
                hasil(i, 2) = "Dimensi tidak valid"
            End If
        Else
            hasil(i, 1) = Empty
            hasil(i, 2) = "Bukan angka"
        End If
    Next i

    ws.Range("E2:F6").Value = hasil
    ws.Range("E2:E6").NumberFormat = "0.000"
End Sub
```

### Struktur lembar `DataVolume`

| Kolom | Isi |
|---|---|
| A | Segmen |
| B | Panjang (m) |
| C | Lebar (m) |
| D | Tinggi (m) |
| E | Volume (m³) |
| F | Status |

Gunakan lima data, termasuk satu nilai nol dan satu teks salah, agar validasi terlihat bekerja.

## 8. Mengapa pemisahan ini penting

`VolumeBalok` tidak mengetahui letak sel mana pun. Karena itu fungsi dapat diuji langsung, dan dapat dipakai ulang oleh macro tabel maupun UserForm pada Modul 7. `HitungTabelVolume` bertanggung jawab atas alur Excel, bukan rumus geometri.

Komentar sebaiknya menjelaskan **alasan, asumsi, atau satuan** — bukan mengulang sintaks:

```vb
' Asumsi latihan: massa jenis air 1000 kg/m³ pada kondisi umum.
Const RHO_AIR_KGM3 As Double = 1000#
```

Nama `x`, `a`, atau `temp` boleh untuk loop kecil, tetapi `debit_m3s` jauh lebih jelas untuk hasil teknik.

## 9. Debugging di Visual Basic Editor

| Alat | Cara | Fungsi |
|---|---|---|
| Compile | `Debug` → `Compile VBAProject` | menemukan galat sintaks/deklarasi |
| Breakpoint | klik margin atau `F9` | menghentikan program di satu baris |
| Step Into | `F8` | menjalankan satu baris demi satu baris |
| Immediate Window | `Ctrl+G` | mencoba ekspresi dan melihat `Debug.Print` |
| Watch | pilih variabel → `Add Watch` | memantau perubahan nilai |
| Locals Window | `View` → `Locals Window` | melihat semua variabel lokal |

Uji fungsi langsung di Immediate Window tanpa menyiapkan sel:

```text
? VolumeBalok(10, 2, 0.3)
? VolumeBalok(0, 2, 0.3)
? LuasLingkaran(2)
```

Hasil yang diharapkan adalah `6`, `-1`, dan `≈3,1416`.

Menekan `F8` sambil memperhatikan Locals Window adalah versi hidup dari **tabel jejak** pada Modul 4 §7 — dengan komputer yang mengisi kolomnya.

## 10. Penanganan galat terarah (demonstrasi dosen)

Gunakan *error handler* untuk galat yang memang mungkin terjadi, dan berikan pesan yang membantu.

```vb
Sub HitungKecepatanDariSel()
    On Error GoTo TanganiGalat

    Dim ws As Worksheet
    Dim debit_m3s As Double
    Dim luas_m2 As Double

    Set ws = ThisWorkbook.Worksheets("Kecepatan")

    If Not IsNumeric(ws.Range("B2").Value) Or _
       Not IsNumeric(ws.Range("B3").Value) Then
        MsgBox "Debit dan luas harus berupa angka.", vbExclamation
        Exit Sub
    End If

    debit_m3s = CDbl(ws.Range("B2").Value)
    luas_m2 = CDbl(ws.Range("B3").Value)

    If debit_m3s < 0 Or luas_m2 <= 0 Then
        MsgBox "Debit tidak boleh negatif dan luas harus > 0.", vbExclamation
        Exit Sub
    End If

    ws.Range("B5").Value = debit_m3s / luas_m2
    Exit Sub

TanganiGalat:
    MsgBox "Program berhenti: " & Err.Description, vbCritical
End Sub
```

Perhatikan bahwa validasi eksplisit tetap dikerjakan lebih dulu; *error handler* hanya jaring pengaman terakhir.

> Jangan memakai `On Error Resume Next` untuk menyembunyikan semua galat. Jika terpaksa dipakai pada satu operasi khusus, segera periksa `Err.Number`, lalu aktifkan kembali penanganan normal.

## 11. Pengujian dengan nilai acuan

Fungsi pembanding berikut memeriksa nilai aktual terhadap nilai harapan dengan toleransi — konsekuensi langsung dari representasi bilangan pecahan yang dibahas pada Modul 1 §4.

```vb
Private Sub PeriksaMendekati( _
    ByVal namaUji As String, _
    ByVal aktual As Double, _
    ByVal harapan As Double, _
    ByVal toleransi As Double)

    If Abs(aktual - harapan) <= toleransi Then
        Debug.Print "LULUS - " & namaUji
    Else
        Debug.Print "GAGAL - " & namaUji & _
                    "; harapan=" & harapan & _
                    "; aktual=" & aktual
    End If
End Sub

Sub UjiLuasLingkaran()
    Const PI As Double = 3.14159265358979

    PeriksaMendekati "diameter 2 m", LuasLingkaran(2), PI, 0.000001
    PeriksaMendekati "diameter 1 m", LuasLingkaran(1), PI / 4, 0.000001
    PeriksaMendekati "input nol", LuasLingkaran(0), -1, 0.000001
End Sub
```

### Pola Arrange–Act–Assert

1. **Arrange:** siapkan input dan nilai harapan.
2. **Act:** panggil fungsi.
3. **Assert:** bandingkan hasil aktual dengan harapan.

Uji setidaknya kasus normal, batas, tidak valid, dan nilai ekstrem yang masih masuk akal. Ini adalah tabel uji Modul 3 §7 yang kini dijalankan oleh program dan dapat diulang kapan saja tanpa biaya.

## 12. Praktik mandiri — fungsi volume pipa

Buat fungsi:

```vb
Public Function VolumePipa( _
    ByVal diameter_m As Double, _
    ByVal panjang_m As Double) As Double
```

Ketentuan:

- gunakan `LuasLingkaran` di dalam fungsi, **bukan menyalin rumusnya**;
- kembalikan `-1` jika diameter atau panjang tidak positif;
- dokumentasikan parameter dan satuan; dan
- buat sedikitnya empat pengujian otomatis.

Nilai acuan untuk diameter `1 m` dan panjang `10 m` adalah sekitar **`7,8539816 m³`**.

### Tantangan debugging

Temukan dan perbaiki kesalahan pada kode berikut, lalu tuliskan jenis kesalahannya:

```vb
Function Tekanan(ByVal gaya_kN As Double, ByVal luas_m2 As Double) As Double
    If luas_m2 < 0 Then
        Tekanan = gaya_kN * luas_m2
    Else
        Tekanan = 0
    End If
End Function
```

### Pengayaan/PR

Buat lembar `Hujan` dengan 30 nilai hujan harian pada `A2:A31`. Tulis macro yang membaca data ke array, menolak nilai negatif atau nonnumerik, lalu menghitung jumlah hari valid, total, rata-rata, dan maksimum **tanpa fungsi worksheet**. Untuk data uji `10, 0, 25, 5, 10`: jumlah data 5, total 50 mm, rata-rata 10 mm, maksimum 25 mm.

## Kuis berdampak — 3 soal

### 1. Prediksi

Variabel `total_m3` dideklarasikan di dalam `Sub A`. Apakah `Sub B` dapat langsung membaca nilainya? Apa perubahan desain yang lebih aman daripada menjadikannya public?

### 2. Praktik perbaikan

Sebuah uji ditulis sebagai `If LuasLingkaran(2) = 3.14159265358979 Then`. Uji itu gagal walaupun fungsinya benar. Jelaskan penyebabnya dan tulis versi yang benar.

### 3. Jelaskan

Mengapa rumus volume sebaiknya berada di `Function VolumeBalok`, sedangkan pembacaan dan penulisan tabel berada di `Sub HitungTabelVolume`? Berikan satu manfaat konkret untuk pengujian dan satu untuk Modul 7.

<details>
<summary>Kunci dan indikator pemahaman</summary>

1. Tidak. Variabel local hanya tersedia di prosedurnya. Desain yang lebih aman adalah mengirim nilai sebagai parameter atau mengembalikannya dari `Function`.
2. Perbandingan bilangan pecahan dengan `=` hampir selalu gagal karena selisih pembulatan sangat kecil. Gunakan toleransi: `If Abs(LuasLingkaran(2) - PI) <= 0.000001 Then`.
3. Pemisahan tanggung jawab membuat rumus tidak bergantung pada worksheet. Untuk pengujian: fungsi dapat dipanggil langsung di Immediate Window dengan input dan hasil harapan, tanpa menyiapkan sel. Untuk Modul 7: UserForm dapat memakai ulang `VolumeBalok` yang sama tanpa menyalin rumusnya.

</details>

## Checklist

- [ ] Semua variabel dideklarasikan dan bertipe sesuai.
- [ ] Input diperiksa sebelum dikonversi.
- [ ] Saya membedakan local dan public, serta `ByVal` dan `ByRef`.
- [ ] Batas array ditangani dengan `LBound`/`UBound`.
- [ ] Setiap fungsi mempunyai satu tanggung jawab utama.
- [ ] Parameter dan satuan terdokumentasi.
- [ ] Saya menjalankan `Compile VBAProject`.
- [ ] Fungsi diuji dengan kasus normal, batas, dan tidak valid.
- [ ] Perbandingan bilangan pecahan memakai toleransi, bukan `=`.
- [ ] Lima baris tetap diproses meskipun satu baris salah.

## Ringkasan

Modularitas memisahkan rumus teknik dari alur spreadsheet, sehingga rumus dapat diuji sendiri dan dipakai ulang. Struktur data mengatur sekumpulan nilai agar diproses konsisten. Debugging menemukan penyebab kesalahan, sedangkan pengujian memberi bukti berulang bahwa fungsi bekerja pada kasus yang dirancang — termasuk kasus yang seharusnya ditolak.

## Bacaan lanjut

1. Michael Alexander & Dick Kusleika, *Excel 2019 Power Programming with VBA*, Wiley, 2019.
2. Bill Jelen & Tracy Syrstad, *Microsoft Excel VBA and Macros (Office 2021 and Microsoft 365)*, Microsoft Press, 2022.
3. Steve McConnell, *Code Complete*, 2nd ed., Microsoft Press, 2004.
4. Brian W. Kernighan & Rob Pike, *The Practice of Programming*, Addison-Wesley, 1999.
5. E. Joseph Billo, *Excel for Scientists and Engineers: Numerical Methods*, Wiley, 2007.

[← Modul 5](05-vba-macro-percabangan-perulangan.md) · [Daftar modul](README.md) · [Modul 7 →](07-userform-otomasi-dan-praktik-integratif.md)
