# Modul 1 — Pengantar Komputer, OS, dan Pemrograman Teknik Sipil

## Capaian pembelajaran

Mahasiswa mampu:

- menjelaskan tujuan, ruang lingkup, dan metode belajar mata kuliah;
- menggambarkan hubungan input, CPU, memori, penyimpanan, output, dan OS;
- membaca bilangan biner sederhana;
- membedakan program, script, dan perangkat lunak aplikasi; serta
- memberi alasan kapan memakai aplikasi yang tersedia dan kapan membuat otomasi sendiri.

## Alur 100 menit

| Menit | Kegiatan |
|---:|---|
| 0–10 | Kontrak belajar, tujuan, dan contoh persoalan sipil |
| 10–22 | Sejarah komputer secara ringkas |
| 22–38 | Komponen komputer dan cara kerja OS |
| 38–53 | Bilangan desimal, biner, bit, dan byte |
| 53–68 | Program komputer dalam teknik sipil |
| 68–82 | Demo VBA Excel dan VBA AutoCAD |
| 82–95 | Kuis tiga soal |
| 95–100 | Pemeriksaan dan *exit ticket* |

**Keluaran minimum:** satu diagram cara kerja komputer, satu keputusan *pakai–adaptasi–buat*, dan hasil pengamatan dua script sederhana.

## 1. Tujuan dan cara belajar

Mata kuliah ini melatih cara berpikir komputasional: memecah masalah, menuliskan langkah, menerjemahkannya menjadi instruksi, dan memeriksa hasil. Keterampilan utamanya adalah:

- **dekomposisi:** memecah masalah besar menjadi langkah kecil;
- **abstraksi:** memilih data dan asumsi yang relevan;
- **pengenalan pola:** melihat hitungan yang berulang;
- **algoritma:** menyusun urutan langkah yang jelas;
- **debugging:** mencari penyebab hasil salah; dan
- **validasi:** membandingkan hasil program dengan kasus acuan.

Metode kelas adalah *predict–run–explain*: prediksi hasil, jalankan, lalu jelaskan perbedaan antara prediksi dan hasil aktual.

## 2. Sejarah komputer dalam empat lompatan

| Tahap | Perubahan penting | Dampak |
|---|---|---|
| Mekanik | alat hitung dan mesin mekanik | operasi aritmetika dibantu alat |
| Elektronik awal | tabung vakum dan program mesin | hitungan kompleks dapat diotomasi |
| Transistor dan IC | perangkat lebih kecil dan andal | komputer masuk organisasi dan laboratorium |
| Mikroprosesor–jaringan | PC, internet, cloud, perangkat bergerak | komputasi tersedia hampir di setiap pekerjaan |

Inti yang tidak berubah adalah: komputer menerima data, menjalankan instruksi, menyimpan keadaan, dan menghasilkan keluaran.

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

Contoh OS adalah Windows, Linux, macOS, Android, dan iOS. Excel dan AutoCAD adalah aplikasi yang berjalan di atas OS; VBA adalah bahasa/lingkungan otomasi yang bekerja di dalam aplikasi tertentu.

## 4. Bilangan dan proses dalam komputer

Satu **bit** mempunyai dua keadaan: `0` atau `1`. Delapan bit membentuk satu **byte**. Nilai biner dibaca dengan bobot pangkat dua.

| Posisi | 3 | 2 | 1 | 0 |
|---|---:|---:|---:|---:|
| Bobot | 8 | 4 | 2 | 1 |
| Digit untuk 13 | 1 | 1 | 0 | 1 |

Jadi `1101₂ = 8 + 4 + 0 + 1 = 13₁₀`.

Komputer juga merepresentasikan teks, gambar, dan bilangan pecahan sebagai pola bit. Karena bilangan pecahan tertentu tidak dapat direpresentasikan secara tepat, hitungan komputer dapat memiliki selisih pembulatan sangat kecil. Dalam teknik, gunakan toleransi ketika membandingkan hasil pecahan.

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

Pilihan ketiga sering paling baik: gunakan aplikasi yang ada lalu otomasi bagian repetitif dengan script.

## 6. Program dan script

Program adalah kumpulan instruksi yang menyelesaikan tugas. Script biasanya lebih kecil, dijalankan di dalam lingkungan tertentu, dan mengotomasi alur kerja. Batas keduanya tidak selalu tegas.

### Demo A — VBA Excel

```vb
Option Explicit

Sub DemoExcel()
    Range("A1").Value = "Panjang (m)"
    Range("B1").Value = 10
    Range("A2").Value = "Lebar (m)"
    Range("B2").Value = 2
    Range("A3").Value = "Luas (m²)"
    Range("B3").Value = Range("B1").Value * Range("B2").Value
End Sub
```

Amati urutan: menulis label, menulis input, membaca input, menghitung, lalu menulis output.

### Demo B — VBA AutoCAD (opsional)

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

## Kuis berdampak — 3 soal

### 1. Prediksi

Tanpa kalkulator, ubah `10110₂` ke desimal. Tunjukkan bobot setiap digit.

### 2. Praktik keputusan

Anda harus menghitung volume dan biaya untuk 500 segmen saluran setiap minggu. Pilih: memakai aplikasi yang tersedia, membuat aplikasi besar, atau membuat script Excel kecil. Berikan dua alasan dan satu risiko.

### 3. Jelaskan

Pada demo Excel, jelaskan apa yang berperan sebagai input, proses, output, aplikasi, bahasa script, dan OS.

<details>
<summary>Kunci dan indikator pemahaman</summary>

**Kunci dan indikator pemahaman**

1. `10110₂ = 16 + 0 + 4 + 2 + 0 = 22₁₀`.
2. Jawaban yang masuk akal adalah script Excel kecil karena alurnya khusus dan berulang. Risiko: rumus/satuan salah atau data lama tertimpa; mitigasinya pengujian, validasi, dan salinan data.
3. Input: B1/B2; proses: perkalian; output: B3; aplikasi: Excel; bahasa: VBA; OS: misalnya Windows yang mengelola aplikasi dan perangkat.

</details>

## Checklist

- [ ] Saya dapat menggambar hubungan CPU, memori, penyimpanan, dan OS.
- [ ] Saya dapat mengubah bilangan biner empat/lima digit ke desimal.
- [ ] Saya dapat menjelaskan perbedaan aplikasi dan script.
- [ ] Saya dapat memberi alasan kapan otomasi kecil layak dibuat.

## Bacaan lanjut

1. J. Glenn Brookshear & Dennis Brylow, *Computer Science: An Overview*, 13th ed., Pearson, 2019.
2. Andrew S. Tanenbaum & Herbert Bos, *Modern Operating Systems*, 4th ed., Pearson, 2014.
3. Abraham Silberschatz, Peter B. Galvin & Greg Gagne, *Operating System Concepts*, 10th ed., Wiley, 2018.
4. Ronald W. Larsen, *Engineering with Excel*, 5th ed., Pearson, 2017.

[← Daftar modul](README.md) · [Modul 2 →](02-dasar-excel.md)
