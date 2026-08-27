# Modul 4 — Algoritma: Representasi, Penelusuran, dan Verifikasi

## Capaian pembelajaran

Mahasiswa mampu:

- memecah masalah teknik menjadi input–proses–output;
- menjelaskan urutan, percabangan, dan perulangan sebagai elemen algoritma;
- menulis algoritma dalam bahasa natural, pseudocode, dan flowchart;
- menelusuri algoritma menggunakan tabel jejak;
- menjelaskan hubungan algoritma dengan program komputer; serta
- memverifikasi hasil algoritma terhadap satu hitungan manual.

## Alur 100 menit

| Menit | Kegiatan | Bagian |
|---:|---|---|
| 0–10 | Permainan menyusun langkah yang diacak | §1 |
| 10–22 | Definisi algoritma, IPO, dan kasus debit | §1–2 |
| 22–34 | Urutan, keputusan, dan perulangan | §3 |
| 34–48 | Narasi, pseudocode, dan flowchart | §4–5 |
| 48–60 | Flowchart kasus volume galian | §6 |
| 60–72 | Tabel jejak sebagai alat debugging | §7 |
| 72–80 | Menerjemahkan algoritma ke Excel/VBA | §8 |
| 80–90 | Praktik tekanan pondasi | §9 |
| 90–97 | Latihan singkat tiga soal | Latihan |
| 97–100 | Pemeriksaan dan checklist | Checklist |

**Keluaran minimum:** satu tabel IPO, pseudocode, flowchart, dan tabel jejak untuk satu kasus teknik, ditutup dengan satu hitungan manual sebagai pembanding.

## 1. Pemecahan masalah menjadi langkah

Contoh persoalan: menghitung total volume galian dari beberapa segmen berbentuk balok.

Sebelum memikirkan kode, jawab:

1. data apa yang tersedia?
2. data apa yang belum tersedia?
3. rumus dan asumsi apa yang digunakan?
4. hasil apa yang dibutuhkan?
5. bagaimana mengetahui hasilnya benar?

### Input–proses–output

| Komponen | Isi |
|---|---|
| Input | panjang, lebar, kedalaman setiap segmen |
| Proses | validasi dimensi, hitung volume per segmen, jumlahkan |
| Output | volume tiap segmen dan total volume |
| Validasi | satu hitungan manual dan pemeriksaan semua dimensi positif |

Pertanyaan kelima adalah yang paling sering dilewati. Algoritma yang tidak menyebutkan cara memeriksa hasilnya belum selesai.

## 2. Definisi algoritma

Algoritma adalah urutan langkah yang terbatas, jelas, dan dapat dijalankan untuk mengubah input menjadi output. Algoritma yang baik memiliki:

- awal dan akhir;
- input dan output yang jelas;
- langkah yang tidak ambigu;
- urutan yang dapat dilaksanakan;
- kondisi berhenti; dan
- cara memeriksa hasil.

Algoritma tidak bergantung pada bahasa. Algoritma yang sama dapat diterjemahkan ke formula Excel, VBA, Python, atau dikerjakan manual.

### Kasus kedua: debit aliran

Untuk penampang dengan kecepatan rata-rata seragam:

```text
Q = A × v
```

dengan `Q` = debit (m³/s), `A` = luas penampang (m²), dan `v` = kecepatan (m/s). Untuk saluran persegi panjang, `A = b × h`, sehingga `Q = b × h × v`.

| Komponen | Isi |
|---|---|
| Input | lebar `b` (m), kedalaman `h` (m), kecepatan `v` (m/s) |
| Proses | hitung `A = b × h`, lalu `Q = A × v` |
| Output | luas `A` (m²) dan debit `Q` (m³/s) |
| Validasi | satuan hasil harus m³/s; hitung satu kasus manual |

## 3. Tiga struktur dasar

### Urutan

Langkah dijalankan dari atas ke bawah.

```text
baca panjang
baca lebar
baca kedalaman
hitung volume
tampilkan volume
```

### Percabangan

Program memilih langkah berdasarkan kondisi.

```text
JIKA semua dimensi > 0
  hitung volume
LAINNYA
  tampilkan "Input tidak valid"
```

### Perulangan

Langkah yang sama diterapkan pada banyak data.

```text
UNTUK setiap segmen
  periksa dimensi
  hitung volume
  tambahkan ke total
SELESAI UNTUK
```

Semua program terstruktur dapat dibangun dari kombinasi ketiga pola ini. Modul 5 menerjemahkan ketiganya menjadi sintaks VBA.

## 4. Tiga cara menuliskan algoritma

Untuk kasus debit di §2:

### Narasi

Baca lebar, kedalaman, dan kecepatan. Hitung luas penampang. Kalikan luas dengan kecepatan. Tampilkan luas dan debit.

Narasi mudah ditulis tetapi mudah pula menjadi ambigu ketika langkahnya bertambah.

### Pseudocode

```text
MULAI
  BACA lebar_m, kedalaman_m, kecepatan_ms
  luas_m2 ← lebar_m × kedalaman_m
  debit_m3s ← luas_m2 × kecepatan_ms
  TULIS luas_m2, debit_m3s
SELESAI
```

Pseudocode tidak terikat aturan sintaks bahasa mana pun, tetapi cukup tegas untuk diterjemahkan langsung menjadi kode.

### Flowchart

```mermaid
flowchart TD
    A([Mulai]) --> B[/Baca b, h, v/]
    B --> C[Hitung A = b × h]
    C --> D[Hitung Q = A × v]
    D --> E[/Tampilkan A dan Q/]
    E --> F([Selesai])
```

Flowchart paling membantu ketika ada percabangan atau perulangan, karena jalur alternatif terlihat sebagai cabang. GitHub merender diagram Mermaid ini langsung pada halaman Markdown.

## 5. Pseudocode lengkap dengan percabangan dan perulangan

```text
MULAI
  total_m3 ← 0
  BACA jumlah_segmen

  UNTUK i ← 1 SAMPAI jumlah_segmen
    BACA panjang_m, lebar_m, kedalaman_m

    JIKA panjang_m > 0 DAN lebar_m > 0 DAN kedalaman_m > 0
      volume_m3 ← panjang_m × lebar_m × kedalaman_m
      total_m3 ← total_m3 + volume_m3
      TULIS volume_m3
    LAINNYA
      TULIS "Input tidak valid"
    SELESAI JIKA
  SELESAI UNTUK

  TULIS total_m3
SELESAI
```

`total_m3 ← 0` disebut **inisialisasi**. Tanpa nilai awal, penjumlahan berulang tidak memiliki titik mulai yang jelas.

## 6. Flowchart kasus volume galian

```mermaid
flowchart TD
    A([Mulai]) --> B[total = 0]
    B --> C[/Baca data segmen/]
    C --> D{Masih ada segmen?}
    D -- Tidak --> H[/Tampilkan total/]
    H --> I([Selesai])
    D -- Ya --> E{Semua dimensi > 0?}
    E -- Ya --> F[Hitung volume dan tambahkan ke total]
    E -- Tidak --> G[Tandai input tidak valid]
    F --> C
    G --> C
```

Simbol oval menunjukkan awal/akhir, jajar genjang input/output, persegi panjang proses, dan belah ketupat keputusan.

Perhatikan bahwa cabang "tidak valid" **kembali ke alur utama**, bukan menghentikan program. Satu data buruk tidak boleh membatalkan pemrosesan data lainnya — prinsip ini diterapkan sebagai kode pada Modul 5.

## 7. Tabel jejak

Gunakan tiga segmen:

| Segmen | p | l | d | Volume | Total setelah langkah |
|---|---:|---:|---:|---:|---:|
| S1 | 10 | 2 | 1 | 20 | 20 |
| S2 | 5 | 2 | 1 | 10 | 30 |
| S3 | 4 | 0 | 1 | tidak valid | 30 |

Tabel jejak memperlihatkan perubahan variabel setelah setiap iterasi. Ini adalah alat debugging yang dapat digunakan **bahkan sebelum kode ditulis** — dan alat yang sama dipakai lagi pada Modul 6 dengan bantuan debugger VBA.

## 8. Hubungan algoritma dan program

Baris pseudocode:

```text
volume_m3 ← panjang_m × lebar_m × kedalaman_m
```

dapat menjadi formula Excel:

```excel
=B2*C2*D2
```

atau perintah VBA:

```vb
volume_m3 = panjang_m * lebar_m * kedalaman_m
```

Bahasa mengubah cara penulisan, bukan logika dasarnya. Karena itu kesalahan algoritma tidak dapat diperbaiki dengan mengganti bahasa.

## 9. Praktik mandiri — tekanan merata pada pondasi

Buat algoritma lengkap untuk menghitung tekanan rata-rata pada pondasi:

```text
q = P / (B × L)
```

dengan `P` dalam kN, `B` dan `L` dalam m, sehingga `q` dalam kN/m² atau kPa.

Produk yang dikumpulkan:

1. tabel input–proses–output;
2. pseudocode, termasuk penanganan `B` atau `L` bernilai nol;
3. flowchart;
4. tabel jejak untuk tiga kasus; dan
5. satu hitungan manual sebagai pembanding.

Gunakan data awal `P = 900 kN`, `B = 2 m`, `L = 3 m`. **Hasil acuan adalah `150 kPa`.**

### Pengayaan/PR

Tambahkan input beban dalam ton-gaya, lalu konversikan ke kN dengan asumsi `1 tf = 9,80665 kN`. Tuliskan asumsi konversi pada lembar kerja sebagai sel berlabel, jangan menanamkannya di dalam rumus.

## Latihan singkat — 3 soal

### 1. Prediksi dan urutkan

Urutkan langkah berikut: `tampilkan hasil`, `validasi pembagi`, `baca gaya`, `hitung tegangan = gaya/luas`, `baca luas`. Tambahkan tindakan jika luas sama dengan nol.

### 2. Temukan kesalahan

Pseudocode total volume langsung menjalankan `total ← total + volume`, tetapi tidak pernah memberi nilai awal pada `total`. Prediksi akibatnya dan perbaiki algoritmanya.

### 3. Jelaskan dengan tabel jejak

Untuk data volume `3, 5, -2, 4`, algoritma hanya menjumlahkan nilai positif. Buat tabel jejak `i`, `nilai`, dan `total`, lalu jelaskan mengapa hasil akhirnya 12, bukan 10.

<details>
<summary>Kunci dan indikator pemahaman</summary>

1. Baca gaya → baca luas → validasi luas → jika luas bukan nol hitung tegangan → tampilkan hasil; jika nol tampilkan pesan dan jangan membagi.
2. Total tidak mempunyai keadaan awal yang pasti. Tambahkan `total ← 0` sebelum perulangan.
3. Total berubah `0→3→8→8→12`; `-2` dilewati, bukan ditambahkan. Mahasiswa harus membedakan "melewati data" dan "menambahkan nilai negatif" — keduanya menghasilkan angka berbeda dan hanya salah satu yang sesuai spesifikasi.

</details>

## Checklist

- [ ] Saya menentukan IPO sebelum menulis kode.
- [ ] Algoritma saya memiliki awal, akhir, dan kondisi berhenti.
- [ ] Setiap akumulator diberi nilai awal.
- [ ] Saya membedakan urutan, percabangan, dan perulangan.
- [ ] Flowchart saya sesuai dengan pseudocode-nya.
- [ ] Saya dapat menelusuri perubahan nilai dengan tabel jejak.
- [ ] Ada satu hitungan manual sebagai pembanding hasil.

## Ringkasan

Algoritma menjembatani rumus teknik dan kode. Narasi, pseudocode, dan flowchart adalah tiga cara menuliskan gagasan yang sama dengan tingkat ketegasan berbeda, sedangkan tabel jejak membuktikan bahwa gagasan itu benar-benar berjalan sebagaimana diduga. Memeriksa algoritma sebelum menulis sintaks jauh lebih murah daripada mencari kesalahan setelah kode berjalan.

## Bacaan lanjut

1. Thomas H. Cormen, *Algorithms Unlocked*, MIT Press, 2013.
2. David Harel & Yishai Feldman, *Algorithmics: The Spirit of Computing*, 3rd ed., Addison-Wesley, 2004.
3. Thomas H. Cormen, Charles E. Leiserson, Ronald L. Rivest & Clifford Stein, *Introduction to Algorithms*, 4th ed., MIT Press, 2022.
4. Steven C. Chapra & Raymond P. Canale, *Numerical Methods for Engineers*, 8th ed., McGraw-Hill, 2021.

[← Modul 3](03-keputusan-dan-pengujian-batas.md) · [Daftar modul](README.md) · [Modul 5 →](05-vba-macro-percabangan-perulangan.md)
