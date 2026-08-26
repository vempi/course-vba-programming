# Modul 5 — Pengenalan VBA dan Macro Linier

## Capaian pembelajaran

Mahasiswa mampu:

- mengenali komponen Excel yang diakses VBA;
- membuka Visual Basic Editor dan memahami struktur `Sub`;
- menjelaskan compile, runtime, statement, dan galat secara sederhana;
- merekam operasi spreadsheet menjadi macro;
- membaca serta memodifikasi hasil Macro Recorder; dan
- membuat program linier baru tanpa recorder.

## Alur 100 menit

| Menit | Kegiatan |
|---:|---|
| 0–10 | Menelusuri objek Workbook–Worksheet–Range |
| 10–25 | Visual Basic Editor, module, `Sub`, dan statement |
| 25–42 | Merekam macro operasi Excel |
| 42–57 | Membaca dan membersihkan hasil rekaman |
| 57–75 | Membuat macro linier volume beton |
| 75–84 | Compile, jalankan, dan bedakan jenis galat |
| 84–97 | Kuis tiga soal |
| 97–100 | Simpan `.xlsm` dan checklist |

**Keluaran minimum:** satu macro rekaman yang telah dibersihkan dan satu macro linier buatan sendiri yang membaca input, menghitung, serta menulis hasil.

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

## 3. Compile, runtime, dan tiga jenis galat

- **Compile/check:** VBA memeriksa sintaks dan deklarasi sebelum prosedur dijalankan. Gunakan `Debug` → `Compile VBAProject`.
- **Runtime:** statement dijalankan dan berinteraksi dengan nilai atau objek aktual.
- **Galat sintaks:** aturan penulisan salah.
- **Galat runtime:** kode mulai berjalan lalu gagal, misalnya worksheet tidak ditemukan.
- **Galat logika:** kode berjalan tetapi hasil salah.

Istilah “compiler” dan “interpreter” menjelaskan cara menerjemahkan instruksi. Dalam praktik awal VBA, fokus mahasiswa adalah memisahkan kesalahan yang ditemukan saat compile, saat berjalan, dan melalui pemeriksaan hasil.

## 4. Macro Recorder

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

Recorder berguna untuk menemukan nama objek, property, dan method. Hasilnya adalah titik awal yang perlu dibaca dan dirapikan.

## 5. Program linier

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
    Dim panjang_m As Double
    Dim lebar_m As Double
    Dim tinggi_m As Double
    Dim volume_m3 As Double

    panjang_m = ThisWorkbook.Worksheets("Volume").Range("B2").Value
    lebar_m = ThisWorkbook.Worksheets("Volume").Range("B3").Value
    tinggi_m = ThisWorkbook.Worksheets("Volume").Range("B4").Value

    volume_m3 = panjang_m * lebar_m * tinggi_m

    ThisWorkbook.Worksheets("Volume").Range("B6").Value = volume_m3
    ThisWorkbook.Worksheets("Volume").Range("B6").NumberFormat = "0.000"
End Sub
```

Hasil acuan adalah `6,000 m³`.

## 6. Memodifikasi program

Tambahkan faktor kehilangan 5%:

```vb
Const FAKTOR_KEHILANGAN As Double = 1.05
Dim volumePesan_m3 As Double

volumePesan_m3 = volume_m3 * FAKTOR_KEHILANGAN
ThisWorkbook.Worksheets("Volume").Range("B7").Value = volumePesan_m3
```

Uji program setelah setiap perubahan kecil. Jangan menunggu sampai banyak perubahan menumpuk.

## Kuis berdampak — 3 soal

### 1. Prediksi

Macro recorder menghasilkan `Range("B2").Select` lalu `Selection.Value = 10`. Apa risiko jika worksheet lain sedang aktif? Tulis versi yang tidak bergantung pada seleksi.

### 2. Praktik perbaikan

Kode volume menghasilkan `60 m³`, padahal data 10 m × 2 m × 0,3 m. Gunakan tabel jejak untuk menemukan satu contoh kesalahan logika yang mungkin dan perbaiki.

### 3. Jelaskan

Kelompokkan tiga kejadian: `End Sub` hilang; worksheet `Volume` tidak ada; rumus memakai penjumlahan alih-alih perkalian. Mana galat sintaks/compile, runtime, dan logika?

<details>
<summary>Kunci dan indikator pemahaman</summary>

1. Sel B2 pada sheet aktif dapat berubah. Gunakan `ThisWorkbook.Worksheets("Volume").Range("B2").Value = 10`.
2. Contoh penyebab: tinggi dibaca sebagai `3`, bukan `0,3`, atau satuan cm belum dikonversi. Tabel jejak harus menunjukkan nilai input dan hasil antara.
3. `End Sub` hilang: sintaks/compile; sheet tidak ada: runtime; penjumlahan menggantikan perkalian: logika.

</details>

## Checklist

- [ ] Workbook disimpan sebagai `.xlsm`.
- [ ] Saya dapat menelusuri Workbook–Worksheet–Range.
- [ ] Saya memahami setiap baris penting hasil recorder.
- [ ] Macro linier cocok dengan hitungan manual.

[← Modul 4](04-algoritma.md) · [Daftar modul](README.md) · [Modul 6 →](06-input-output-dan-modularitas.md)
