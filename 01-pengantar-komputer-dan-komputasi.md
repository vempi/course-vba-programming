# Modul 1 — Pengantar Komputer, OS, dan Komputasi Teknik Sipil

## Capaian pembelajaran

Mahasiswa mampu:

- menjelaskan tujuan, ruang lingkup, dan metode belajar mata kuliah;
- menguraikan persoalan teknik menjadi input–proses–output;
- menggambarkan hubungan input, CPU, memori, penyimpanan, output, dan OS;
- membaca bilangan biner sederhana;
- membedakan program, script, dan perangkat lunak aplikasi;
- memberi alasan kapan memakai aplikasi yang tersedia dan kapan membuat otomasi sendiri; serta
- menyiapkan workbook `.xlsm` dan menjalankan macro VBA pertama.

## Alur 100 menit

| Menit | Kegiatan | Bagian |
|---:|---|---|
| 0–8 | Kontrak belajar, tujuan, dan contoh persoalan sipil | §1 |
| 8–18 | Alur masalah → model → algoritma → kode → validasi | §1 |
| 18–28 | Sejarah komputer, komponen, dan cara kerja OS | §2–3 |
| 28–38 | Bilangan desimal, biner, bit, dan byte | §4 |
| 38–48 | Program dalam teknik sipil dan keputusan pakai–adaptasi–buat | §5–6 |
| 48–56 | Menyiapkan Excel, Developer, dan Visual Basic Editor | §7 |
| 56–66 | Demonstrasi macro pertama | §8 |
| 66–78 | Demonstrasi volume pelat dan jejak eksekusi | §9 |
| 78–88 | Praktik mandiri volume sloof + 5% | §11 |
| 88–97 | Kuis tiga soal | Kuis |
| 97–100 | Pemeriksaan dan *exit ticket* | Checklist |

**Keluaran minimum:** satu diagram cara kerja komputer, satu keputusan *pakai–adaptasi–buat*, satu workbook `.xlsm`, dan satu macro yang hasilnya cocok dengan hitungan manual.

Demo AutoCAD (§10) dan pembanding Python (§9) cukup ditunjukkan dosen selama 2–3 menit. Mahasiswa tidak wajib menjalankannya di kelas.

## 1. Dari masalah teknik ke program

Mata kuliah ini melatih cara berpikir komputasional: memecah masalah, menuliskan langkah, menerjemahkannya menjadi instruksi, dan memeriksa hasil. Keterampilan utamanya:

- **dekomposisi:** memecah masalah besar menjadi langkah kecil;
- **abstraksi:** memilih data dan asumsi yang relevan;
- **pengenalan pola:** melihat hitungan yang berulang;
- **algoritma:** menyusun urutan langkah yang jelas;
- **debugging:** mencari penyebab hasil salah; dan
- **validasi:** membandingkan hasil program dengan kasus acuan.

Metode kelas adalah *predict–run–explain*: prediksi hasil, jalankan, lalu jelaskan perbedaan antara prediksi dan hasil aktual.

Komputasi teknik bukan sekadar mengetik rumus. Alur lengkapnya:

```text
masalah nyata → asumsi/model → data input → algoritma → kode
             → hasil → validasi → keputusan teknik
```

Contoh: menghitung volume pelat beton berbentuk balok.

- **Masalah:** berapa volume beton yang dibutuhkan?
- **Asumsi:** pelat berbentuk balok sempurna; kehilangan material belum dihitung.
- **Input:** panjang, lebar, dan tebal dalam meter.
- **Algoritma:** `volume = panjang × lebar × tebal`.
- **Output:** volume dalam m³.
- **Validasi:** hitung satu kasus secara manual dan periksa satuannya.

> **Prinsip yang berlaku sepanjang mata kuliah.** Program yang menghasilkan angka belum tentu benar. Rumus bisa salah, satuan bisa tidak konsisten, atau input bisa tidak masuk akal. Validasi adalah bagian dari pemrograman, bukan pekerjaan tambahan setelah program selesai.

## 2. Sejarah komputer dalam empat lompatan

| Tahap | Perubahan penting | Dampak |
|---|---|---|
| Mekanik | alat hitung dan mesin mekanik | operasi aritmetika dibantu alat |
| Elektronik awal | tabung vakum dan program mesin | hitungan kompleks dapat diotomasi |
| Transistor dan IC | perangkat lebih kecil dan andal | komputer masuk organisasi dan laboratorium |
| Mikroprosesor–jaringan | PC, internet, cloud, perangkat bergerak | komputasi tersedia hampir di setiap pekerjaan |

Inti yang tidak berubah: komputer menerima data, menjalankan instruksi, menyimpan keadaan, dan menghasilkan keluaran.

## 3. Cara kerja komputer dan OS

```text
pengguna/data
     ↓
perangkat input → memori ↔ CPU → perangkat output
                    ↕
               penyimpanan
                    ↕
          Operating System (OS)
```

CPU menjalankan instruksi dasar. Memori menyimpan data yang sedang dipakai. Penyimpanan menjaga data setelah daya dimatikan. OS mengelola perangkat keras, file, memori, proses, pengguna, dan antarmuka agar aplikasi tidak harus mengendalikan perangkat secara langsung.

Contoh OS adalah Windows, Linux, macOS, Android, dan iOS. Excel dan AutoCAD adalah aplikasi yang berjalan di atas OS; VBA adalah bahasa dan lingkungan otomasi yang bekerja **di dalam** aplikasi tertentu.

## 4. Bilangan dan proses dalam komputer

Satu **bit** mempunyai dua keadaan: `0` atau `1`. Delapan bit membentuk satu **byte**. Nilai biner dibaca dengan bobot pangkat dua.

| Posisi | 3 | 2 | 1 | 0 |
|---|---:|---:|---:|---:|
| Bobot | 8 | 4 | 2 | 1 |
| Digit untuk 13 | 1 | 1 | 0 | 1 |

Jadi `1101₂ = 8 + 4 + 0 + 1 = 13₁₀`.

Komputer juga merepresentasikan teks, gambar, dan bilangan pecahan sebagai pola bit. Karena bilangan pecahan tertentu tidak dapat direpresentasikan secara tepat, hitungan komputer dapat memiliki selisih pembulatan sangat kecil. Dalam teknik, **gunakan toleransi ketika membandingkan hasil pecahan** — bukan tanda sama dengan. Konsekuensi praktisnya muncul lagi pada Modul 6 saat menulis pengujian.

## 5. Pemanfaatan program dalam teknik sipil

Contoh pemanfaatan:

- spreadsheet untuk kuantitas, biaya, dan tabulasi data;
- CAD/BIM untuk gambar dan model informasi;
- GIS untuk data spasial;
- perangkat analisis struktur, geoteknik, transportasi, atau hidraulika;
- script untuk membersihkan data, hitungan berulang, pelaporan, dan pemeriksaan mutu.

### Pakai, adaptasi, atau buat?

| Pertanyaan | Cenderung memakai aplikasi tersedia | Cenderung membuat script/aplikasi kecil |
|---|---|---|
| Apakah masalah umum dan sudah teruji? | ya | tidak |
| Apakah alur khusus dan berulang? | kadang | ya |
| Apakah risiko keselamatan tinggi? | gunakan perangkat tervalidasi | script hanya sebagai alat bantu dan wajib diverifikasi |
| Apakah biaya membuat lebih besar dari manfaat? | ya | tidak |

Pilihan ketiga sering paling baik: gunakan aplikasi yang ada, lalu otomasi bagian repetitifnya dengan script.

## 6. Excel, VBA, dan Python — kapan memakai apa

Program adalah kumpulan instruksi yang menyelesaikan tugas. Script biasanya lebih kecil, dijalankan di dalam lingkungan tertentu, dan mengotomasi alur kerja. Batas keduanya tidak selalu tegas.

| Kebutuhan | Excel | VBA | Python |
|---|---|---|---|
| Satu hitungan langsung | `=B2*C2*D2` | bisa, tetapi belum perlu | bisa, tetapi belum perlu |
| Mengulang 1.000 baris | salin rumus | loop/macro | loop atau operasi array |
| Memeriksa input | Data Validation/`IF` | `If...Then` dan pesan | `if` dan exception |
| Membuat tombol proses | Form Control | sangat sesuai | perlu antarmuka lain |
| Analisis data besar/lanjut | terbatas | cukup untuk skala kecil–menengah | sangat sesuai |

Prinsipnya sama: input dibaca, proses dijalankan, hasil ditulis. Perbedaannya terutama pada sintaks dan lingkungan kerja. Mata kuliah ini memakai VBA sebagai bahasa utama karena paling dekat dengan tabel Excel yang sudah dikenal mahasiswa.

## 7. Menyiapkan Excel untuk VBA

1. Buka Excel desktop dan buat workbook kosong.
2. Simpan sebagai **Excel Macro-Enabled Workbook (`.xlsm`)**.
3. Tampilkan tab **Developer** melalui `File` → `Options` → `Customize Ribbon` → centang `Developer`.
4. Tekan `Alt+F11`, lalu pilih `Insert` → `Module`.
5. Pastikan baris pertama kode adalah `Option Explicit`.

`Option Explicit` memaksa setiap variabel dideklarasikan. Ini menangkap salah ketik nama variabel sebelum hasil yang keliru menyebar ke seluruh lembar kerja.

## 8. Demonstrasi 1 — macro pertama

Tempel kode berikut ke module, letakkan kursor di dalam prosedur, lalu tekan `F5`.

```vb
Option Explicit

Sub SapaTeknikSipil()
    MsgBox "Halo! Kita mulai komputasi teknik sipil.", _
           vbInformation, "Pengantar Komputer dan VBA"
End Sub
```

Anatomi singkat:

- `Sub ... End Sub` membatasi sebuah prosedur;
- `MsgBox` menampilkan output sederhana;
- tanda `_` melanjutkan satu perintah ke baris berikutnya; dan
- teks diapit tanda petik ganda.

## 9. Demonstrasi 2 — volume pelat beton

Buat lembar kerja bernama `Volume` dengan isi berikut.

| Sel | Isi | Sel | Isi |
|---|---|---|---|
| A2 | Panjang (m) | B2 | 6 |
| A3 | Lebar (m) | B3 | 4 |
| A4 | Tebal (m) | B4 | 0,15 |
| A6 | Volume (m³) | B6 | *hasil program* |

```vb
Option Explicit

Sub HitungVolumePelat()
    Dim ws As Worksheet
    Dim panjang_m As Double
    Dim lebar_m As Double
    Dim tebal_m As Double
    Dim volume_m3 As Double

    Set ws = ThisWorkbook.Worksheets("Volume")

    panjang_m = ws.Range("B2").Value
    lebar_m = ws.Range("B3").Value
    tebal_m = ws.Range("B4").Value

    volume_m3 = panjang_m * lebar_m * tebal_m

    ws.Range("B6").Value = volume_m3
    ws.Range("B6").NumberFormat = "0.000"
End Sub
```

Hasil acuan untuk data di atas adalah **3,600 m³**.

Perhatikan urutannya: membaca input, menghitung, lalu menulis output — pola input–proses–output yang sama seperti pada §1. Rujukan `ThisWorkbook.Worksheets("Volume")` dipakai sejak awal agar kode tidak bergantung pada lembar yang kebetulan aktif.

### Jejak eksekusi

| Langkah | Variabel/aksi | Nilai |
|---:|---|---:|
| 1 | baca `panjang_m` | 6 |
| 2 | baca `lebar_m` | 4 |
| 3 | baca `tebal_m` | 0,15 |
| 4 | hitung `volume_m3` | 3,6 |
| 5 | tulis ke B6 | 3,6 |

Tabel jejak seperti ini akan dipakai berulang kali sebagai alat debugging, bahkan sebelum kode ditulis.

### Padanan di Python (pengayaan singkat)

```python
panjang_m = 6.0
lebar_m = 4.0
tebal_m = 0.15

volume_m3 = panjang_m * lebar_m * tebal_m
print(f"Volume = {volume_m3:.3f} m³")
```

Sintaks berbeda, alur input–proses–output tetap sama.

## 10. Demonstrasi 3 — VBA AutoCAD (opsional)

Demo ini dijalankan hanya jika lingkungan VBA AutoCAD tersedia. Satuan mengikuti satuan gambar aktif.

```vb
Sub BuatGarisSederhana()
    Dim titikAwal(0 To 2) As Double
    Dim titikAkhir(0 To 2) As Double
    Dim garis As AcadLine

    titikAwal(0) = 0: titikAwal(1) = 0: titikAwal(2) = 0
    titikAkhir(0) = 10: titikAkhir(1) = 5: titikAkhir(2) = 0

    Set garis = ThisDrawing.ModelSpace.AddLine(titikAwal, titikAkhir)
    ZoomAll
End Sub
```

Hierarki `ThisDrawing.ModelSpace.AddLine` mengikuti model objek ActiveX AutoCAD pada [dokumentasi resmi Autodesk](https://help.autodesk.com/cloudhelp/2024/CHS/AutoCAD-ActiveX/files/GUID-D4FF317D-16DA-42D8-8309-8260B7427E55.htm).

Poin demo ini: VBA bukan milik Excel saja. Bahasa yang sama menggerakkan aplikasi berbeda melalui model objek masing-masing.

## 11. Praktik mandiri — volume sloof

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

Kasus batas sengaja dibiarkan terbuka. Modul 3 dan Modul 5 menunjukkan cara program menolak input seperti itu, alih-alih menghasilkan angka nol yang tampak wajar.

## Pertanyaan refleksi

1. Bagian mana yang merupakan model, algoritma, dan kode?
2. Mengapa nilai nol pada panjang perlu diperiksa?
3. Bagaimana memastikan semua input memakai meter, bukan sentimeter?
4. *Exit ticket* — satu kalimat: kapan rumus Excel sudah cukup dan kapan VBA lebih bermanfaat?

## Kuis berdampak — 3 soal

### 1. Prediksi

Tanpa kalkulator, ubah `10110₂` ke desimal. Tunjukkan bobot setiap digit.

### 2. Praktik keputusan

Anda harus menghitung volume dan biaya untuk 500 segmen saluran setiap minggu. Pilih: memakai aplikasi yang tersedia, membuat aplikasi besar, atau membuat script Excel kecil. Berikan dua alasan dan satu risiko.

### 3. Jelaskan

Pada `HitungVolumePelat`, jelaskan apa yang berperan sebagai input, proses, output, aplikasi, bahasa script, dan OS. Lalu sebutkan satu cara memeriksa bahwa hasil `3,600` memang benar.

<details>
<summary>Kunci dan indikator pemahaman</summary>

1. `10110₂ = 16 + 0 + 4 + 2 + 0 = 22₁₀`.
2. Jawaban yang masuk akal adalah script Excel kecil karena alurnya khusus dan berulang. Risiko: rumus/satuan salah atau data lama tertimpa; mitigasinya pengujian, validasi, dan salinan data.
3. Input: B2/B3/B4; proses: perkalian di dalam macro; output: B6; aplikasi: Excel; bahasa: VBA; OS: misalnya Windows yang mengelola aplikasi dan perangkat. Pemeriksaan: hitung `6 × 4 × 0,15` secara manual dan bandingkan, sekaligus pastikan seluruh input memakai meter.

</details>

## Checklist

- [ ] Saya dapat menggambar hubungan CPU, memori, penyimpanan, dan OS.
- [ ] Saya dapat mengubah bilangan biner empat/lima digit ke desimal.
- [ ] Saya dapat menjelaskan perbedaan aplikasi dan script.
- [ ] Saya dapat memberi alasan kapan otomasi kecil layak dibuat.
- [ ] Workbook tersimpan sebagai `.xlsm` dan macro berjalan tanpa pesan galat.
- [ ] Hasil macro cocok dengan hitungan manual.
- [ ] Nama variabel saya menjelaskan makna dan satuan.

## Ringkasan

Komputer menerima data, menjalankan instruksi, dan menghasilkan keluaran; OS mengelola sumber daya agar aplikasi tidak perlu mengurus perangkat keras. Di atas lapisan itu, VBA mengotomasi pekerjaan berulang di dalam Excel. Yang membedakan program teknik dari sekadar kode adalah asumsi yang terbuka, satuan yang konsisten, dan hasil yang diverifikasi.

## Bacaan lanjut

1. J. Glenn Brookshear & Dennis Brylow, *Computer Science: An Overview*, 13th ed., Pearson, 2019.
2. Andrew S. Tanenbaum & Herbert Bos, *Modern Operating Systems*, 4th ed., Pearson, 2014.
3. Abraham Silberschatz, Peter B. Galvin & Greg Gagne, *Operating System Concepts*, 10th ed., Wiley, 2018.
4. Ronald W. Larsen, *Engineering with Excel*, 5th ed., Pearson, 2017.

[← Daftar modul](README.md) · [Modul 2 →](02-excel-fungsi-dan-satuan.md)
