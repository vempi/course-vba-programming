# Modul 7 — UserForm dan Control dalam VBA

## Capaian pembelajaran

Mahasiswa mampu:

- menjelaskan peran UserForm sebagai antarmuka;
- menambahkan dan menamai Control;
- membedakan property, method, dan event;
- menulis event handler tombol;
- memvalidasi input form; serta
- menggunakan kembali Function dari module standar.

## Alur 100 menit

| Menit | Kegiatan |
|---:|---|
| 0–10 | Membandingkan input lewat sel dan form |
| 10–25 | Membuat UserForm dan menambah Control |
| 25–40 | Name, Caption, Value, property, dan event |
| 40–58 | Demo event tombol Hitung |
| 58–76 | Praktik form volume beton |
| 76–86 | Validasi, tombol Bersihkan, dan Tutup |
| 86–97 | Kuis tiga soal dan demonstrasi pasangan |
| 97–100 | Simpan, compile, dan checklist akhir |

**Keluaran minimum:** satu UserForm dengan tiga input, satu output, tombol Hitung, Bersihkan, dan Tutup; input salah menghasilkan pesan yang jelas.

## 1. Kapan menggunakan form?

Input langsung di worksheet sesuai untuk tabel besar dan data yang perlu terlihat bersama. UserForm sesuai untuk:

- memandu pengguna melalui sejumlah input tertentu;
- membatasi urutan dan pilihan;
- mengurangi risiko pengguna mengubah formula; dan
- memberi pesan validasi sebelum data disimpan.

Form tidak otomatis membuat program benar. Rumus, satuan, validasi, dan pengujian tetap diperlukan.

## 2. Membuat UserForm

1. Buka Visual Basic Editor dengan `Alt+F11`.
2. Pilih `Insert` → `UserForm`.
3. Pada Properties Window, ubah `(Name)` menjadi `frmVolume`.
4. Ubah `Caption` menjadi `Kalkulator Volume Beton`.
5. Tambahkan Control dari Toolbox.

### Control yang digunakan

| Jenis | `(Name)` | `Caption`/fungsi |
|---|---|---|
| Label | `lblPanjang` | Panjang (m) |
| TextBox | `txtPanjang` | input panjang |
| Label | `lblLebar` | Lebar (m) |
| TextBox | `txtLebar` | input lebar |
| Label | `lblTinggi` | Tinggi (m) |
| TextBox | `txtTinggi` | input tinggi |
| Label | `lblHasil` | Volume: — |
| CommandButton | `cmdHitung` | Hitung |
| CommandButton | `cmdBersihkan` | Bersihkan |
| CommandButton | `cmdTutup` | Tutup |

Nama seperti `txtPanjang` menjelaskan jenis dan makna Control. Ini lebih mudah dirawat daripada nama bawaan `TextBox1`.

## 3. Property, method, dan event

- **Property** adalah keadaan/atribut, misalnya `Caption`, `Name`, `Value`, dan `Enabled`.
- **Method** adalah tindakan yang diminta, misalnya `SetFocus`, `Show`, dan `Hide`.
- **Event** adalah kejadian yang ditanggapi, misalnya `Click`, `Change`, dan `Initialize`.

```text
pengguna klik tombol
        ↓
event cmdHitung_Click
        ↓
baca property Value dari TextBox
        ↓
panggil Function VolumeBalok
        ↓
ubah property Caption pada lblHasil
```

## 4. Menampilkan form

Pada module standar:

```vb
Option Explicit

Public Sub BukaFormVolume()
    frmVolume.Show
End Sub
```

Jalankan `BukaFormVolume` atau hubungkan macro ini ke tombol pada worksheet.

## 5. Event tombol Hitung

Klik ganda `cmdHitung`, lalu isi event handler:

```vb
Private Sub cmdHitung_Click()
    Dim panjang_m As Double
    Dim lebar_m As Double
    Dim tinggi_m As Double
    Dim volume_m3 As Double

    If Not IsNumeric(Me.txtPanjang.Value) Or _
       Not IsNumeric(Me.txtLebar.Value) Or _
       Not IsNumeric(Me.txtTinggi.Value) Then
        MsgBox "Semua dimensi harus berupa angka.", vbExclamation
        Exit Sub
    End If

    panjang_m = CDbl(Me.txtPanjang.Value)
    lebar_m = CDbl(Me.txtLebar.Value)
    tinggi_m = CDbl(Me.txtTinggi.Value)

    volume_m3 = VolumeBalok(panjang_m, lebar_m, tinggi_m)

    If volume_m3 < 0 Then
        MsgBox "Semua dimensi harus lebih besar dari nol.", vbExclamation
        Exit Sub
    End If

    Me.lblHasil.Caption = "Volume: " & _
                           Format(volume_m3, "0.000") & " m³"
End Sub
```

`Me` menunjuk UserForm tempat kode berada. Fungsi `VolumeBalok` menggunakan fungsi yang sudah dibuat pada Modul 6 di module standar.

## 6. Event Bersihkan dan Tutup

```vb
Private Sub cmdBersihkan_Click()
    Me.txtPanjang.Value = ""
    Me.txtLebar.Value = ""
    Me.txtTinggi.Value = ""
    Me.lblHasil.Caption = "Volume: —"
    Me.txtPanjang.SetFocus
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub
```

Inisialisasi form:

```vb
Private Sub UserForm_Initialize()
    Me.lblHasil.Caption = "Volume: —"
    Me.txtPanjang.Value = ""
    Me.txtLebar.Value = ""
    Me.txtTinggi.Value = ""
End Sub
```

## 7. Praktik dan pengujian

Jalankan kasus berikut:

| Kasus | p | l | t | Hasil yang diharapkan |
|---|---:|---:|---:|---|
| Normal | 10 | 2 | 0,3 | 6,000 m³ |
| Nol | 0 | 2 | 0,3 | pesan dimensi harus > 0 |
| Teks | `sepuluh` | 2 | 0,3 | pesan harus berupa angka |
| Pecahan | 1,5 | 0,2 | 0,4 | 0,120 m³ |

Demo individu selama satu menit:

1. buka form;
2. jalankan kasus normal;
3. masukkan satu input salah;
4. jelaskan event yang berjalan; dan
5. tunjukkan Function yang digunakan ulang.

## 8. Batas desain antarmuka

- selalu tampilkan satuan;
- gunakan urutan tab yang logis;
- jangan mengandalkan warna sebagai satu-satunya pesan;
- berikan pesan yang menyebutkan apa yang salah dan cara memperbaikinya;
- jangan menyimpan hasil sebelum input valid; dan
- sediakan cara membatalkan atau menutup form.

## Kuis berdampak — 3 soal

### 1. Prediksi

Pengguna mengetik `abc` pada `txtPanjang` lalu menekan Hitung. Urutkan statement penting yang berjalan dan jelaskan mengapa `CDbl` tidak dipanggil.

### 2. Praktik perbaikan

Tombol Bersihkan mengosongkan tiga TextBox tetapi hasil lama masih terlihat. Tambahkan satu statement yang diperlukan dan jelaskan property yang berubah.

### 3. Jelaskan

Kelompokkan contoh berikut sebagai property, method, atau event: `Caption`, `SetFocus`, `Click`, `Value`, `Show`, `Initialize`. Lalu jelaskan hubungan ketiganya ketika tombol Hitung ditekan.

<details>
<summary>Kunci dan indikator pemahaman</summary>

1. Event `cmdHitung_Click` berjalan, `IsNumeric` menghasilkan False, pesan muncul, lalu `Exit Sub`. Karena keluar lebih awal, `CDbl` tidak dijalankan dan tidak menghasilkan type mismatch.
2. `Me.lblHasil.Caption = "Volume: —"`; yang berubah adalah property `Caption` milik Label.
3. Property: `Caption`, `Value`; method: `SetFocus`, `Show`; event: `Click`, `Initialize`. Event memicu handler, handler membaca/mengubah property dan dapat memanggil method.

</details>

## Checklist akhir

- [ ] Semua Control mempunyai nama yang bermakna.
- [ ] Saya membedakan property, method, dan event.
- [ ] Validasi dilakukan sebelum `CDbl` dan sebelum hitungan.
- [ ] Form menggunakan Function dari module standar.
- [ ] Saya dapat mendemonstrasikan kasus normal dan kasus salah.

## Penutup

UserForm menghubungkan pengguna dengan algoritma. Antarmuka yang baik memandu input, tetapi kualitas hasil tetap bergantung pada fungsi yang benar, validasi yang jelas, dan pengujian yang dapat dibuktikan.

## Bacaan lanjut

1. Michael Alexander & Dick Kusleika, *Excel 2019 Power Programming with VBA*, Wiley, 2019.
2. Bill Jelen & Tracy Syrstad, *Microsoft Excel VBA and Macros (Office 2021 and Microsoft 365)*, Microsoft Press, 2022.
3. Steve Krug, *Don’t Make Me Think, Revisited: A Common Sense Approach to Web Usability*, 3rd ed., New Riders, 2014.
4. Ben Shneiderman dkk., *Designing the User Interface: Strategies for Effective Human-Computer Interaction*, 6th ed., Pearson, 2016.

[← Modul 6](06-input-output-dan-modularitas.md) · [Daftar modul](README.md)
