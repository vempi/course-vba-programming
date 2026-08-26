# Modul 5 — Fungsi, Modularitas, Debugging, dan Pengujian

## Capaian pembelajaran

Mahasiswa mampu:

- memecah program menjadi `Sub` dan `Function` yang terfokus;
- menulis dokumentasi singkat tentang parameter, hasil, dan satuan;
- menelusuri kode dengan breakpoint, *step execution*, dan Immediate Window;
- menangani galat yang dapat diperkirakan; dan
- menyusun pengujian sederhana dengan hasil yang diketahui.

## Alur 100 menit

| Menit | Kegiatan | Bagian modul |
|---:|---|---|
| 0–10 | Memecah program panjang menjadi tugas kecil | §1 |
| 10–30 | Demonstrasi `Function`, parameter, dan satuan | §2–4 |
| 30–45 | Debugging dengan compile, breakpoint, `F8`, dan Immediate Window | §5–6 |
| 45–55 | Validasi dan penanganan galat terarah | §7 |
| 55–80 | Praktik membuat `VolumePipa` | §10 |
| 80–93 | Menulis dan menjalankan tiga uji | §8–9 |
| 93–100 | Tantangan debugging singkat dan checklist | §11 dan checklist |

**Keluaran minimum:** satu fungsi `VolumePipa`, dokumentasi parameter/satuan, dan tiga pengujian yang mencakup kasus normal, batas/tidak valid, serta nilai acuan. Pembahasan `ByRef` dan *error handler* lengkap dapat dilanjutkan sebagai bahan baca setelah kelas.

### Satu benang merah

Gunakan kasus pipa dari awal sampai akhir: buat fungsi luas, pakai fungsi itu untuk volume, sengaja masukkan kesalahan, telusuri dengan debugger, lalu buktikan perbaikannya dengan pengujian. Dengan demikian mahasiswa tidak perlu berpindah-pindah konteks kasus.

## 1. Modularitas

Program yang baik dibagi menjadi bagian kecil dengan satu tanggung jawab.

```text
Sub utama
├── membaca dan memvalidasi input
├── memanggil fungsi hitungan
└── menulis hasil
```

`Sub` menjalankan tindakan, sedangkan `Function` mengembalikan nilai. Fungsi kecil mudah diuji tanpa harus menjalankan seluruh workbook.

## 2. Membuat fungsi teknik

Contoh fungsi luas lingkaran:

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

Fungsi dapat dipanggil dari VBA:

```vb
Sub DemoLuasPipa()
    Dim luas_m2 As Double
    luas_m2 = LuasLingkaran(1.2)
    MsgBox Format(luas_m2, "0.000") & " m²"
End Sub
```

Karena `Public Function` berada di module standar, fungsi juga dapat dipakai pada sel Excel: `=LuasLingkaran(B2)`.

## 3. `ByVal`, `ByRef`, dan ruang lingkup (konsep ringkas)

- `ByVal` memberikan salinan nilai; perubahan di fungsi tidak mengubah variabel pemanggil.
- `ByRef` memberikan referensi; fungsi/prosedur dapat mengubah variabel pemanggil.
- variabel lokal hanya hidup di dalam prosedur;
- konstanta module dapat dibagikan, tetapi terlalu banyak variabel global membuat program sulit ditelusuri.

Gunakan `ByVal` sebagai pilihan awal untuk parameter hitungan.

## 4. Dokumentasi yang berguna

Komentar sebaiknya menjelaskan alasan, asumsi, atau satuan—bukan mengulang sintaks.

```vb
' Asumsi latihan: massa jenis air 1000 kg/m³ pada kondisi umum.
Const RHO_AIR_KGM3 As Double = 1000#
```

Nama `x`, `a`, atau `temp` boleh untuk loop kecil, tetapi `debit_m3s` jauh lebih jelas untuk hasil teknik.

## 5. Debugging di Visual Basic Editor

Alat utama:

| Alat | Cara | Fungsi |
|---|---|---|
| Compile | `Debug` → `Compile VBAProject` | menemukan galat sintaks/deklarasi |
| Breakpoint | klik margin atau `F9` | menghentikan program di satu baris |
| Step Into | `F8` | menjalankan satu baris demi satu baris |
| Immediate Window | `Ctrl+G` | mencoba ekspresi dan melihat `Debug.Print` |
| Watch | pilih variabel → `Add Watch` | memantau perubahan nilai |
| Locals Window | `View` → `Locals Window` | melihat semua variabel lokal |

Contoh penggunaan Immediate Window saat program berhenti:

```text
? diameter_m
? LuasLingkaran(2)
```

## 6. Tiga kelompok galat

1. **Galat sintaks/compile:** kode tidak dapat dijalankan, misalnya `End If` hilang.
2. **Galat runtime:** kode mulai berjalan lalu berhenti, misalnya membagi dengan nol.
3. **Galat logika:** kode berjalan tetapi hasil salah, misalnya diameter dipakai sebagai jari-jari. Ini paling berbahaya karena tidak selalu ada pesan.

## 7. Penanganan galat terarah (demonstrasi dosen)

Gunakan *error handler* untuk galat yang memang mungkin terjadi dan berikan pesan yang membantu.

```vb
Sub HitungKecepatanDariSel()
    On Error GoTo TanganiGalat

    Dim ws As Worksheet
    Dim debit_m3s As Double
    Dim luas_m2 As Double
    Dim kecepatan_ms As Double

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

    kecepatan_ms = debit_m3s / luas_m2
    ws.Range("B5").Value = kecepatan_ms
    Exit Sub

TanganiGalat:
    MsgBox "Program berhenti: " & Err.Description, vbCritical
End Sub
```

Jangan memakai `On Error Resume Next` untuk menyembunyikan semua galat. Jika terpaksa digunakan pada satu operasi khusus, segera periksa `Err.Number`, lalu aktifkan kembali penanganan normal.

## 8. Pengujian dengan nilai acuan

Fungsi pembanding berikut memeriksa nilai aktual terhadap nilai harapan dengan toleransi.

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

Gunakan toleransi untuk bilangan pecahan karena representasi bilangan komputer tidak selalu persis sama.

## 9. Pola Arrange–Act–Assert

Setiap uji dapat dibaca sebagai:

1. **Arrange:** siapkan input dan nilai harapan.
2. **Act:** panggil fungsi.
3. **Assert:** bandingkan hasil aktual dengan harapan.

Uji setidaknya kasus normal, batas, tidak valid, dan nilai ekstrem yang masih masuk akal.

## 10. Praktik mandiri — fungsi volume pipa

Buat fungsi:

```vb
Public Function VolumePipa( _
    ByVal diameter_m As Double, _
    ByVal panjang_m As Double) As Double
```

Ketentuan:

- gunakan `LuasLingkaran` di dalam fungsi, bukan menyalin rumusnya;
- kembalikan `-1` jika diameter atau panjang tidak positif;
- dokumentasikan parameter dan satuan; dan
- buat sedikitnya empat pengujian otomatis.

Nilai acuan untuk diameter `1 m` dan panjang `10 m` adalah sekitar `7,8539816 m³`.

## 11. Tantangan debugging dan *exit ticket*

Temukan dan perbaiki kesalahan pada kode berikut:

```vb
Function Tekanan(ByVal gaya_kN As Double, ByVal luas_m2 As Double) As Double
    If luas_m2 < 0 Then
        Tekanan = gaya_kN * luas_m2
    Else
        Tekanan = 0
    End If
End Function
```

Tuliskan jenis setiap kesalahan: sintaks, runtime, atau logika. Di kelas, cukup perbaiki kode dan tulis satu uji. Dua uji tambahan dapat diselesaikan sebagai pengayaan/PR.

## Checklist

- [ ] Setiap fungsi mempunyai satu tanggung jawab utama.
- [ ] Parameter dan satuan terdokumentasi.
- [ ] Saya menjalankan `Compile VBAProject`.
- [ ] Galat input ditangani sebelum operasi berbahaya.
- [ ] Fungsi diuji dengan kasus normal, batas, dan tidak valid.

## Ringkasan

Modularitas membuat kode lebih mudah dibaca, digunakan ulang, dan diuji. Debugging menemukan penyebab kesalahan, sedangkan pengujian memberi bukti berulang bahwa fungsi bekerja pada kasus yang dirancang.

[← Modul 4](04-struktur-data.md) · [Daftar modul](README.md) · [Modul 6 →](06-otomasi-perhitungan-teknik-sipil.md)
