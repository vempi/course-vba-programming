# Modul 6 — Input–Output, Variabel, Array, Function, dan Subroutine

## Capaian pembelajaran

Mahasiswa mampu:

- membaca dan menulis data melalui `Cells`, `Range`, `InputBox`, dan `MsgBox`;
- mendeklarasikan variabel dengan tipe dan ruang lingkup yang sesuai;
- menggunakan array untuk sekumpulan data;
- membedakan `Function` dan `Sub`; serta
- menyusun program pengolahan tabel yang modular.

## Alur 100 menit

| Menit | Kegiatan |
|---:|---|
| 0–10 | Mengidentifikasi input–proses–output macro sebelumnya |
| 10–25 | `Cells`, `Range`, `InputBox`, dan `MsgBox` |
| 25–40 | Tipe variabel serta ruang lingkup local/public |
| 40–55 | Array satu dan dua dimensi |
| 55–70 | Perbedaan `Function` dan `Sub` |
| 70–85 | Praktik pengolahan tabel volume |
| 85–97 | Kuis tiga soal |
| 97–100 | Compile dan checklist |

**Keluaran minimum:** satu fungsi volume dan satu `Sub` yang membaca sedikitnya lima baris, memprosesnya melalui array, lalu menulis hasil ke Excel.

## 1. Input dan output sederhana

### `Range` dan `Cells`

```vb
Dim nilai As Double

nilai = Range("B2").Value
Cells(2, 3).Value = nilai * 2
```

`Range("B2")` mudah dibaca untuk alamat tetap. `Cells(baris, kolom)` cocok ketika nomor baris atau kolom berubah di dalam loop.

Gunakan worksheet yang eksplisit:

```vb
Dim ws As Worksheet
Set ws = ThisWorkbook.Worksheets("DataVolume")

nilai = ws.Range("B2").Value
ws.Cells(2, "F").Value = nilai
```

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

`InputBox` biasa mengembalikan teks. Periksa dengan `IsNumeric` sebelum mengubah teks menjadi angka menggunakan `CDbl`.

## 2. Variabel dan tipe data

```vb
Dim panjang_m As Double
Dim jumlahSegmen As Long
Dim namaSegmen As String
Dim dataValid As Boolean
Dim tanggalUkur As Date
```

Gunakan:

- `Double` untuk besaran teknik pecahan;
- `Long` untuk jumlah data dan indeks;
- `String` untuk nama/kode;
- `Boolean` untuk benar/salah; dan
- `Date` untuk tanggal.

`Option Explicit` memaksa deklarasi variabel dan membantu menemukan salah ketik nama.

## 3. Ruang lingkup local dan public

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

Gunakan variabel local sebagai pilihan awal. Variabel public berguna ketika keadaan benar-benar perlu dibagi antarmodule/form, tetapi nilainya lebih sulit ditelusuri dan dapat berubah dari banyak tempat.

## 4. Array

Array menyimpan beberapa nilai sejenis dengan satu nama.

```vb
Sub DemoArray()
    Dim volume_m3(1 To 3) As Double
    Dim i As Long
    Dim total_m3 As Double

    volume_m3(1) = 2.5
    volume_m3(2) = 3.5
    volume_m3(3) = 4

    For i = LBound(volume_m3) To UBound(volume_m3)
        total_m3 = total_m3 + volume_m3(i)
    Next i

    MsgBox "Total = " & total_m3 & " m³"
End Sub
```

Blok Excel langsung menjadi array 2D bertipe `Variant`:

```vb
Dim data As Variant
data = ws.Range("B2:D6").Value

' data(1,1) = B2; data(1,2) = C2; data(1,3) = D2
```

## 5. Function dan Subroutine

`Function` menerima parameter dan mengembalikan satu nilai:

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
```

`Sub` mengatur tindakan seperti membaca tabel, memanggil fungsi, dan menulis hasil:

```vb
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

Gunakan lima data, termasuk satu nilai nol dan satu teks salah, agar validasi terlihat.

## 6. Mengapa modular?

Fungsi `VolumeBalok` tidak mengetahui letak sel. Karena itu fungsi mudah diuji dan dapat digunakan oleh macro tabel maupun UserForm pada Modul 7. `HitungTabelVolume` bertanggung jawab atas alur Excel, bukan rumus geometri.

Uji fungsi di Immediate Window:

```text
? VolumeBalok(10, 2, 0.3)
? VolumeBalok(0, 2, 0.3)
```

Hasil yang diharapkan adalah `6` dan `-1`.

## Kuis berdampak — 3 soal

### 1. Prediksi

Variabel `total_m3` dideklarasikan di dalam `Sub A`. Apakah `Sub B` dapat langsung membaca nilainya? Apa perubahan desain yang lebih aman daripada menjadikannya public?

### 2. Praktik perbaikan

Data ketiga berisi teks `dua`. Tentukan bagian program yang mencegah `CDbl("dua")` dijalankan, lalu jelaskan apa yang ditulis pada hasil.

### 3. Jelaskan

Mengapa rumus volume sebaiknya berada di `Function VolumeBalok`, sedangkan pembacaan dan penulisan tabel berada di `Sub HitungTabelVolume`? Berikan satu manfaat untuk pengujian.

<details>
<summary>Kunci dan indikator pemahaman</summary>

1. Tidak. Variabel local hanya tersedia di prosedurnya. Desain yang lebih aman adalah mengirim nilai sebagai parameter atau mengembalikannya dari Function.
2. `IsNumeric` diperiksa sebelum `CDbl`; hasil volume dikosongkan dan status menjadi `Bukan angka`.
3. Pemisahan tanggung jawab membuat rumus tidak bergantung pada worksheet. Fungsi dapat diuji langsung dengan input dan hasil harapan tanpa menyiapkan sel.

</details>

## Checklist

- [ ] Semua variabel dideklarasikan dan bertipe sesuai.
- [ ] Input diperiksa sebelum dikonversi.
- [ ] Saya membedakan local dan public.
- [ ] Saya membedakan tugas Function dan Sub.
- [ ] Lima baris tetap diproses meskipun satu baris salah.

## Bacaan lanjut

1. Michael Alexander & Dick Kusleika, *Excel 2019 Power Programming with VBA*, Wiley, 2019.
2. Bill Jelen & Tracy Syrstad, *Microsoft Excel VBA and Macros (Office 2021 and Microsoft 365)*, Microsoft Press, 2022.
3. Steve McConnell, *Code Complete*, 2nd ed., Microsoft Press, 2004.
4. Brian W. Kernighan & Rob Pike, *The Practice of Programming*, Addison-Wesley, 1999.
5. E. Joseph Billo, *Excel for Scientists and Engineers: Numerical Methods*, Wiley, 2007.

[← Modul 5](05-vba-dan-macro-linier.md) · [Daftar modul](README.md) · [Modul 7 →](07-userform-dan-control.md)
