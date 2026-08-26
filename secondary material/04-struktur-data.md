# Modul 4 — Struktur Data

## Capaian pembelajaran

Mahasiswa mampu:

- memilih struktur data sesuai bentuk masalah;
- menggunakan array satu dan dua dimensi di VBA;
- menggunakan `Collection` dan `Dictionary` untuk data berkelompok;
- membaca serta menulis blok data Excel secara efisien; dan
- mengenali padanan `list`, `tuple`, `dictionary`, dan NumPy array di Python.

## Alur 100 menit

| Menit | Kegiatan | Bagian modul |
|---:|---|---|
| 0–10 | Mengelompokkan contoh data: tunggal, deret, tabel, kunci–nilai | §1 |
| 10–25 | Konsep array dan indeks | §2 |
| 25–40 | Demonstrasi array hujan dan penelusuran loop | §2 |
| 40–50 | Demonstrasi cepat array 2D dari blok Excel | §4 |
| 50–80 | Praktik ringkasan lima lalu 30 data hujan | §8 |
| 80–90 | Peta padanan VBA–Python | §7 |
| 90–97 | Uji kasus kecil | §8 |
| 97–100 | Checklist dan *exit ticket* | Checklist |

**Keluaran minimum:** mahasiswa dapat membaca deret angka ke array, menghitung total/rata-rata/maksimum, dan menjelaskan padanan konsepnya di Python. `Collection`, `Dictionary`, serta penyelesaian matriks NumPy cukup dikenalkan secara visual; implementasinya menjadi pengayaan.

### Batas materi inti

Fokus praktik kelas adalah **array satu dimensi**. Array dinamis dan 2D didemonstrasikan dosen. `Collection`, `Dictionary`, `tuple`, dan NumPy tetap tersedia agar materi lengkap, tetapi tidak semuanya harus diketik mahasiswa dalam satu pertemuan.

## 1. Mengapa struktur data diperlukan?

Satu variabel menyimpan satu nilai. Data teknik biasanya berupa deret curah hujan, koordinat titik, matriks kekakuan, daftar material, atau pasangan kode–nilai. Struktur data membantu menyimpan nilai-nilai itu sebagai satu kesatuan.

| Bentuk data | VBA | Python | Contoh teknik sipil |
|---|---|---|---|
| Deret berurutan, ukuran tetap | array | `list`/`tuple` | hujan 12 bulan |
| Tabel/baris-kolom | array 2D atau `Range` | NumPy array | elevasi beberapa titik |
| Daftar objek dinamis | `Collection` | `list` | nama stasiun |
| Pasangan kunci–nilai | `Dictionary` | `dict` | kode material–berat jenis |

## 2. Array satu dimensi

Array cocok jika elemen sejenis diakses dengan indeks.

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

`LBound` dan `UBound` membuat loop mengikuti batas array sehingga kode lebih mudah diubah.

## 3. Array dinamis (demonstrasi dosen)

Jika ukuran baru diketahui saat program berjalan, gunakan `ReDim`.

```vb
Sub BacaDataKeArray()
    Dim ws As Worksheet
    Dim dataHujan() As Double
    Dim jumlahData As Long
    Dim i As Long

    Set ws = ThisWorkbook.Worksheets("Hujan")
    jumlahData = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row - 1

    If jumlahData <= 0 Then
        MsgBox "Tidak ada data pada kolom A.", vbExclamation
        Exit Sub
    End If

    ReDim dataHujan(1 To jumlahData)

    For i = 1 To jumlahData
        dataHujan(i) = CDbl(ws.Cells(i + 1, "A").Value)
    Next i
End Sub
```

Jika data berisi teks atau sel kosong, tambahkan pemeriksaan `IsNumeric` sebelum `CDbl`.

## 4. Array dua dimensi dan blok Excel (demonstrasi dosen)

Satu `Range.Value` dapat dipindahkan ke array 2D sekaligus. Cara ini umumnya lebih cepat daripada membaca sel satu per satu.

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

## 5. `Collection` (pengayaan)

`Collection` menyimpan daftar yang ukurannya dapat bertambah.

```vb
Sub DemoCollection()
    Dim namaStasiun As New Collection
    Dim item As Variant

    namaStasiun.Add "Stasiun A"
    namaStasiun.Add "Stasiun B"
    namaStasiun.Add "Stasiun C"

    For Each item In namaStasiun
        Debug.Print item
    Next item
End Sub
```

Tekan `Ctrl+G` di Visual Basic Editor untuk melihat **Immediate Window** dan hasil `Debug.Print`.

## 6. `Dictionary` (pengayaan)

Dictionary menyimpan pasangan kunci–nilai. Dengan *late binding* berikut, pengguna tidak perlu menambahkan referensi library secara manual.

```vb
Sub DemoDictionaryMaterial()
    Dim beratJenis As Object
    Dim kode As String

    Set beratJenis = CreateObject("Scripting.Dictionary")
    beratJenis.Add "BETON", 24#
    beratJenis.Add "BAJA", 78.5
    beratJenis.Add "AIR", 9.81

    kode = UCase(Trim(Range("B2").Value))

    If beratJenis.Exists(kode) Then
        Range("B3").Value = beratJenis(kode)
        Range("C3").Value = "kN/m³"
    Else
        Range("B3").ClearContents
        Range("C3").Value = "Kode tidak ditemukan"
    End If
End Sub
```

`Scripting.Dictionary` tersedia pada Excel Windows. Untuk kompatibilitas lintas platform, array tabel atau `Collection` dapat menjadi alternatif.

## 7. Padanan minor di Python (pengenalan 10 menit)

```python
# list: dapat diubah
hujan_mm = [12.5, 0.0, 31.2, 8.4, 17.9]
hujan_mm.append(10.0)

# tuple: biasanya dipakai untuk data yang tidak ingin diubah
titik_xy = (430125.2, 9145678.4)

# dictionary: pasangan kunci-nilai
berat_jenis = {"BETON": 24.0, "BAJA": 78.5, "AIR": 9.81}
print(berat_jenis["BETON"])
```

NumPy menyediakan operasi vektor/matriks yang ringkas:

```python
import numpy as np

hujan_mm = np.array([12.5, 0.0, 31.2, 8.4, 17.9])
print(hujan_mm.mean())
print(hujan_mm.max())

matriks = np.array([[2.0, -1.0], [-1.0, 2.0]])
vektor = np.array([10.0, 5.0])
solusi = np.linalg.solve(matriks, vektor)
print(solusi)
```

NumPy sangat bermanfaat untuk analisis matriks besar. Dalam VBA, pengolahan serupa memerlukan loop atau fungsi matriks Excel.

## 8. Praktik mandiri — ringkasan curah hujan

Buat lembar `Hujan` dengan 30 nilai hujan harian pada `A2:A31`. Tulis macro yang:

1. membaca data ke array;
2. menolak nilai negatif atau nonnumerik;
3. menghitung jumlah hari valid;
4. menghitung total, rata-rata, dan maksimum tanpa fungsi worksheet; dan
5. menulis ringkasan pada `D2:E5`.

Gunakan algoritma maksimum berikut:

```text
maksimum ← elemen valid pertama
UNTUK setiap nilai berikutnya
  JIKA nilai > maksimum MAKA maksimum ← nilai
SELESAI
```

### Kasus uji kecil

Gunakan data `10, 0, 25, 5, 10`. Hasil yang diharapkan:

- jumlah data = 5;
- total = 50 mm;
- rata-rata = 10 mm; dan
- maksimum = 25 mm.

## 9. Tantangan pengayaan/PR

Buat dictionary untuk menyimpan tiga stasiun dan elevasinya. Pengguna memasukkan kode stasiun pada satu sel; program menampilkan elevasi atau pesan bahwa kode tidak ditemukan.

## Checklist

- [ ] Saya memilih struktur data berdasarkan bentuk masalah, bukan sekadar kebiasaan.
- [ ] Batas array ditangani dengan `LBound`/`UBound`.
- [ ] Data diperiksa sebelum dikonversi menjadi angka.
- [ ] Saya dapat menjelaskan perbedaan array, `Collection`, dan `Dictionary`.
- [ ] Saya mengenali padanan struktur data dasar di Python.

## Ringkasan

Struktur data mengatur sekumpulan nilai agar dapat diproses secara konsisten. Array unggul untuk data berindeks dan tabel; `Collection` untuk daftar dinamis; `Dictionary` untuk pencarian berdasarkan kunci; dan NumPy untuk operasi vektor/matriks di Python.

[← Modul 3](03-struktur-kontrol.md) · [Daftar modul](README.md) · [Modul 5 →](05-fungsi-debugging-pengujian.md)
