# Modul 6 — Otomasi Perhitungan Teknik Sipil

## Capaian pembelajaran

Mahasiswa mampu:

- membuat fungsi konversi satuan yang eksplisit;
- menerjemahkan persamaan teknik menjadi fungsi VBA;
- melakukan interpolasi linear dengan pemeriksaan batas;
- mengotomasi hitungan untuk banyak baris data; dan
- membandingkan implementasi VBA dengan Python secara minor.

## Alur 100 menit

| Menit | Kegiatan | Bagian modul |
|---:|---|---|
| 0–10 | Audit satuan pada contoh input | §1 |
| 10–25 | Demonstrasi fungsi konversi dan satu uji | §1 |
| 25–45 | Menurunkan persamaan Manning menjadi variabel antara | §2 |
| 45–60 | Menjalankan fungsi pada kasus acuan | §2 |
| 60–82 | Praktik mengotomasi lima alternatif | §4 dan §6 |
| 82–92 | Verifikasi manual alternatif pertama | §2 dan §6 |
| 92–97 | Demonstrasi interpolasi linear | §3 |
| 97–100 | Checklist dan *exit ticket* | Checklist |

**Keluaran minimum:** fungsi konversi, fungsi debit Manning, tabel lima alternatif, dan satu verifikasi manual. Python hanya ditunjukkan sebagai padanan sintaks. Mahasiswa tidak wajib mengetik Python atau membuat interpolasi tabel otomatis pada pertemuan ini.

### Batas materi inti

Benang merah pertemuan adalah **input bersatuan → fungsi Manning → loop lima alternatif → verifikasi**. Interpolasi diperkenalkan melalui satu contoh hitungan. Implementasi Python dan pencarian otomatis dua titik pengapit menjadi pengayaan.

## 1. Konversi satuan sebagai fungsi

Konversi yang berulang sebaiknya dibuat sebagai fungsi agar tidak ada faktor yang ditulis berbeda di banyak tempat.

```vb
Option Explicit

Public Function MmKeM(ByVal panjang_mm As Double) As Double
    MmKeM = panjang_mm / 1000#
End Function

Public Function LiterPerDetikKeM3s(ByVal debit_Ls As Double) As Double
    LiterPerDetikKeM3s = debit_Ls / 1000#
End Function

Public Function DerajatKeRadian(ByVal sudut_derajat As Double) As Double
    Const PI As Double = 3.14159265358979
    DerajatKeRadian = sudut_derajat * PI / 180#
End Function
```

Uji cepat:

| Input | Harapan |
|---|---:|
| `MmKeM(2500)` | `2,5 m` |
| `LiterPerDetikKeM3s(500)` | `0,5 m³/s` |
| `DerajatKeRadian(180)` | `π rad` |

## 2. Menerjemahkan persamaan: Manning

Untuk latihan, debit aliran seragam dihitung dengan persamaan Manning SI:

```text
Q = (1/n) × A × R^(2/3) × S^(1/2)
```

dengan `n` = koefisien Manning, `A` = luas basah (m²), `R` = jari-jari hidraulik (m), dan `S` = kemiringan energi (m/m).

Untuk saluran persegi panjang dengan lebar dasar `b` dan kedalaman air `y`:

```text
A = b × y
P = b + 2y
R = A/P
```

Implementasi fungsi:

```vb
' Menghitung debit Manning untuk saluran persegi panjang.
' Semua panjang dalam m; S dalam m/m; hasil dalam m³/s.
' Mengembalikan -1 jika input tidak valid.
Public Function DebitManningPersegi( _
    ByVal lebar_m As Double, _
    ByVal kedalaman_m As Double, _
    ByVal koefManning As Double, _
    ByVal kemiringan As Double) As Double

    Dim luas_m2 As Double
    Dim kelilingBasah_m As Double
    Dim radiusHidraulik_m As Double

    If lebar_m <= 0 Or kedalaman_m <= 0 Or _
       koefManning <= 0 Or kemiringan < 0 Then
        DebitManningPersegi = -1
        Exit Function
    End If

    luas_m2 = lebar_m * kedalaman_m
    kelilingBasah_m = lebar_m + 2# * kedalaman_m
    radiusHidraulik_m = luas_m2 / kelilingBasah_m

    DebitManningPersegi = (1# / koefManning) * luas_m2 * _
                           radiusHidraulik_m ^ (2# / 3#) * _
                           kemiringan ^ 0.5
End Function
```

Tulisan `2# / 3#` menegaskan operasi pecahan bertipe `Double`.

### Kasus acuan

Untuk `b = 2 m`, `y = 1 m`, `n = 0,015`, dan `S = 0,001`:

- `A = 2 m²`;
- `P = 4 m`;
- `R = 0,5 m`; dan
- `Q ≈ 2,6561 m³/s`.

Persamaan ini adalah contoh pembelajaran. Penggunaan untuk desain memerlukan penetapan parameter, kondisi aliran, geometri, dan standar yang tepat.

## 3. Interpolasi linear (demonstrasi ringkas)

Jika diketahui `(x1, y1)` dan `(x2, y2)`, nilai pada `x` di antaranya adalah:

```text
y = y1 + (x - x1) × (y2 - y1) / (x2 - x1)
```

```vb
Public Function InterpolasiLinear( _
    ByVal x As Double, _
    ByVal x1 As Double, ByVal y1 As Double, _
    ByVal x2 As Double, ByVal y2 As Double) As Variant

    If x2 = x1 Then
        InterpolasiLinear = CVErr(xlErrDiv0)
    ElseIf x < WorksheetFunction.Min(x1, x2) Or _
           x > WorksheetFunction.Max(x1, x2) Then
        InterpolasiLinear = CVErr(xlErrNA)
    Else
        InterpolasiLinear = y1 + (x - x1) * _
                            (y2 - y1) / (x2 - x1)
    End If
End Function
```

Fungsi mengembalikan galat Excel jika dua titik memiliki `x` sama atau jika diminta melakukan ekstrapolasi. Kebijakan ini membuat batas penggunaan terlihat jelas.

### Uji interpolasi

Jika `(x1, y1) = (10, 100)` dan `(x2, y2) = (20, 160)`, maka:

- pada `x = 10`, `y = 100`;
- pada `x = 15`, `y = 130`; dan
- pada `x = 20`, `y = 160`.

## 4. Otomasi hitungan berulang

Buat tabel pada lembar `Manning`:

| Kolom | Isi |
|---|---|
| A | Alternatif |
| B | Lebar `b` (m) |
| C | Kedalaman `y` (m) |
| D | Manning `n` |
| E | Kemiringan `S` (m/m) |
| F | Debit `Q` (m³/s) |
| G | Status |

```vb
Sub HitungAlternatifManning()
    Dim ws As Worksheet
    Dim baris As Long
    Dim barisTerakhir As Long
    Dim hasilQ As Double

    Set ws = ThisWorkbook.Worksheets("Manning")
    barisTerakhir = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For baris = 2 To barisTerakhir
        If IsNumeric(ws.Cells(baris, "B").Value) And _
           IsNumeric(ws.Cells(baris, "C").Value) And _
           IsNumeric(ws.Cells(baris, "D").Value) And _
           IsNumeric(ws.Cells(baris, "E").Value) Then

            hasilQ = DebitManningPersegi( _
                CDbl(ws.Cells(baris, "B").Value), _
                CDbl(ws.Cells(baris, "C").Value), _
                CDbl(ws.Cells(baris, "D").Value), _
                CDbl(ws.Cells(baris, "E").Value))

            If hasilQ >= 0 Then
                ws.Cells(baris, "F").Value = hasilQ
                ws.Cells(baris, "G").Value = "OK"
            Else
                ws.Cells(baris, "F").ClearContents
                ws.Cells(baris, "G").Value = "Input tidak valid"
            End If
        Else
            ws.Cells(baris, "F").ClearContents
            ws.Cells(baris, "G").Value = "Bukan angka"
        End If
    Next baris

    ws.Range("F2:F" & barisTerakhir).NumberFormat = "0.0000"
End Sub
```

## 5. Pembanding minor dengan Python (pengayaan)

```python
def debit_manning_persegi(b, y, n, s):
    if b <= 0 or y <= 0 or n <= 0 or s < 0:
        raise ValueError("Input di luar rentang yang diizinkan")
    luas = b * y
    keliling_basah = b + 2 * y
    radius_hidraulik = luas / keliling_basah
    return (1 / n) * luas * radius_hidraulik ** (2 / 3) * s ** 0.5

q = debit_manning_persegi(2.0, 1.0, 0.015, 0.001)
print(f"Q = {q:.4f} m³/s")
```

Untuk banyak alternatif, data dapat disimpan dalam list atau NumPy array. Walau sintaks lebih ringkas, fungsi, validasi, satuan, dan pengujian tetap diperlukan.

## 6. Praktik mandiri

1. Buat tabel sedikitnya lima alternatif saluran.
2. Jalankan macro hitungan berulang.
3. Verifikasi alternatif pertama secara manual.
4. Tambahkan kolom kecepatan `v = Q/A`.
5. Tandai input tidak valid tanpa menghentikan pemrosesan baris lainnya.

Gunakan salah satu baris acuan yang telah diberikan agar implementasi dapat diperiksa.

## 7. Tantangan pengayaan/PR

Buat fungsi interpolasi yang membaca tabel dua kolom di Excel dan otomatis mencari dua titik yang mengapit nilai `x`. Tentukan dengan jelas apa yang terjadi jika data tidak urut atau `x` berada di luar tabel.

## Checklist

- [ ] Konversi satuan dibuat eksplisit dan diuji.
- [ ] Persamaan dipecah menjadi variabel antara yang bermakna.
- [ ] Fungsi menolak domain input yang tidak valid.
- [ ] Interpolasi membedakan interpolasi dan ekstrapolasi.
- [ ] Satu hasil program diverifikasi langkah demi langkah secara manual.

## Ringkasan

Otomasi bukan hanya mempercepat hitungan berulang. Fungsi konversi, fungsi persamaan, pemeriksaan domain, dan verifikasi membuat proses dapat ditelusuri serta mengurangi risiko kesalahan yang berulang.

[← Modul 5](05-fungsi-debugging-pengujian.md) · [Daftar modul](README.md) · [Modul 7 →](07-praktik-integratif-pra-uts.md)
