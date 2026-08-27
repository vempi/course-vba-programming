# -*- coding: utf-8 -*-
"""
Build PowerPoint decks for the Indonesian 7-meeting module:
Pengantar Komputer, Excel, dan VBA.

Run:
    uv run --with python-pptx --with pillow python build_powerpoints.py
"""
import os
import shutil
import sys

sys.path.insert(0, r"D:\OneDrive\Bahan-Kuliah\_Slide-Template-UGM")

from pptx import Presentation
from pptx.util import Inches
from pptx.enum.text import PP_ALIGN

from mn_template_common import (
    pick_source,
    delete_all_slides,
    title_slide,
    agenda_slide,
    divider_slide,
    content_slide,
    closing_slide,
    add_note_box,
    add_table,
    build_and_save,
    NAVY,
    RED,
    DGRAY,
    LGRAY,
    PALE_GOLD,
)


HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = r"D:\OneDrive\Bahan-Kuliah\_Slide-Template-UGM\MN1-Pengantar.pptx"
OUT_DIR = os.path.join(HERE, "00-Slide-Kuliah")
os.makedirs(OUT_DIR, exist_ok=True)

COURSE = "Algoritma dan Pemrograman Komputer"
SUBTITLE = "Excel dan VBA"
LECTURER = "Departemen Teknik Sipil dan Lingkungan FT UGM"
DATE_LABEL = "Agustus 2026"


def bullet(text, size=None, level=0, kind="dot", color=None, bold=False):
    item = {"text": text, "bullet": kind, "level": level}
    if size:
        item["size"] = size
    if color:
        item["color"] = color
    if bold:
        item["bold"] = True
    return item


def plain(text, size=None, color=None, align=None, bold=False):
    item = {"text": text, "bullet": None}
    if size:
        item["size"] = size
    if color:
        item["color"] = color
    if align:
        item["align"] = align
    if bold:
        item["bold"] = True
    return item


def add_schedule(slide, rows):
    add_table(
        slide,
        left=Inches(0.70),
        top=Inches(1.47),
        width=Inches(8.55),
        height=Inches(2.95),
        headers=["Menit", "Kegiatan", "Produk kecil"],
        rows=rows,
        col_widths=[Inches(1.15), Inches(4.15), Inches(3.25)],
        header_size=12,
        body_size=10.5,
        left_cols={1, 2},
        highlight_rows=[len(rows) - 1],
        highlight_fill=PALE_GOLD,
    )


def schedule_slide(prs, title, rows):
    slide = content_slide(prs, "Alur 100 Menit", [
        plain(title, size=15, color=NAVY, bold=True),
    ], base_size=15)
    add_schedule(slide, rows)
    return slide


def exercise_slide(prs, questions, title="Latihan Singkat"):
    slide = content_slide(prs, title, [
        plain("Kerjakan mandiri 6 menit, lalu diskusi cepat 7 menit.", size=15, color=NAVY, bold=True),
        bullet(questions[0], kind="num", size=14.5),
        bullet(questions[1], kind="num", size=14.5),
        bullet(questions[2], kind="num", size=14.5),
    ], base_size=14.5)
    return slide


def code_slide(prs, title, code_lines, note=None):
    slide = content_slide(prs, title, [
        plain("Baca dari atas ke bawah: data masuk, proses, lalu hasil keluar.", size=13.5, color=NAVY, bold=True),
    ], base_size=13.5)
    add_note_box(
        slide,
        left=Inches(0.65),
        top=Inches(1.52),
        width=Inches(8.70),
        height=Inches(3.00),
        fill=LGRAY,
        line=NAVY,
        base_size=10.5,
        items=[plain(line, size=10.5, color=DGRAY) for line in code_lines],
    )
    if note:
        add_note_box(
            slide,
            left=Inches(0.65),
            top=Inches(4.58),
            width=Inches(8.70),
            height=Inches(0.47),
            fill=PALE_GOLD,
            line=NAVY,
            base_size=10,
            items=[plain(note, size=10, color=DGRAY, bold=True)],
        )
    return slide


DECKS = [
    {
        "file": "01-pengantar-komputer-os.pptx",
        "meeting": "Pertemuan 1",
        "topic": "Pengantar Komputer dan OS",
        "agenda": [
            "Tujuan, ruang lingkup, dan cara belajar",
            "Sejarah komputer dan cara kerja OS",
            "Bilangan biner, bit, byte",
            "Program komputer dalam teknik sipil",
            "Demo script: Excel VBA dan AutoCAD VBA",
        ],
        "outcomes": [
            "Menjelaskan tujuan MK dan pola belajar predict-run-explain.",
            "Menggambar hubungan input, CPU, memori, storage, output, dan OS.",
            "Mengonversi bilangan biner sederhana ke desimal.",
            "Membedakan aplikasi siap pakai, program, dan script otomasi.",
        ],
        "schedule": [
            ["0-10", "Kontrak belajar dan contoh kasus sipil", "Pertanyaan awal"],
            ["10-22", "Sejarah komputer ringkas", "Empat lompatan"],
            ["22-38", "Komponen komputer dan OS", "Diagram kerja"],
            ["38-53", "Biner, bit, byte, pembulatan", "Konversi 1101"],
            ["53-68", "Software vs script di teknik sipil", "Keputusan pakai-adaptasi-buat"],
            ["68-82", "Demo VBA Excel dan AutoCAD", "Amati input-proses-output"],
            ["82-100", "Latihan, pembahasan, exit ticket", "3 jawaban inti"],
        ],
        "sections": [
            ("Cara Belajar di MK Ini", [
                bullet("Tujuan utama: berpikir komputasional untuk masalah teknik sipil."),
                bullet("Siklus kelas: **predict** hasil, **run** contoh, lalu **explain** penyebabnya."),
                bullet("Yang dinilai bukan hafalan sintaks, tetapi kemampuan memecah masalah dan memeriksa hasil."),
                bullet("Output minimum tiap pertemuan: contoh kecil yang dapat dijalankan sendiri."),
            ]),
            ("Komputer dan OS", [
                bullet("CPU menjalankan instruksi; memori menyimpan data yang sedang aktif."),
                bullet("Storage menyimpan file; perangkat input-output menghubungkan pengguna dengan sistem."),
                bullet("OS mengatur file, proses, memori, pengguna, perangkat, dan antarmuka."),
                bullet("Excel dan AutoCAD adalah aplikasi; VBA adalah bahasa otomasi di dalam aplikasi."),
            ]),
            ("Bilangan di Komputer", [
                bullet("Satu bit hanya punya dua keadaan: 0 atau 1; delapan bit membentuk satu byte."),
                bullet("Biner dibaca memakai bobot 1, 2, 4, 8, 16, dan seterusnya."),
                bullet("Contoh: 1101 = 8 + 4 + 0 + 1 = 13."),
                bullet("Bilangan pecahan dapat memiliki galat pembulatan kecil; di teknik gunakan toleransi."),
            ]),
            ("Software vs Script", [
                bullet("Gunakan software tervalidasi untuk analisis berisiko tinggi dan masalah standar."),
                bullet("Buat script kecil untuk tugas yang berulang, lokal, dan mudah diverifikasi."),
                bullet("Pilihan praktis: pakai aplikasi yang ada, lalu otomasi bagian yang repetitif."),
                bullet("Risiko utama: salah satuan, salah rumus, dan data lama tertimpa."),
            ]),
        ],
        "code": [
            "Option Explicit",
            "",
            "Sub DemoExcel()",
            "    Range(\"A1\").Value = \"Panjang (m)\"",
            "    Range(\"B1\").Value = 10",
            "    Range(\"A2\").Value = \"Lebar (m)\"",
            "    Range(\"B2\").Value = 2",
            "    Range(\"A3\").Value = \"Luas (m2)\"",
            "    Range(\"B3\").Value = Range(\"B1\").Value * Range(\"B2\").Value",
            "End Sub",
        ],
        "activity": [
            bullet("Ubah nilai panjang dan lebar, lalu prediksi sel mana yang berubah."),
            bullet("Tandai input, proses, dan output pada spreadsheet."),
            bullet("Diskusikan: kapan contoh ini cukup menjadi formula Excel, kapan layak jadi macro?"),
        ],
        "exercise": [
            "Jika pengguna mengisi B1 dan B2 lalu menjalankan DemoExcel, urutkan alur input, Excel, VBA, CPU/memori, worksheet, dan output.",
            "Untuk 500 segmen saluran per minggu, pilih software, aplikasi besar, atau script Excel kecil. Beri dua alasan dan satu risiko.",
            "Pada demo Excel, sebutkan input, proses, output, aplikasi, bahasa script, dan OS.",
        ],
    },
    {
        "file": "02-dasar-excel.pptx",
        "meeting": "Pertemuan 2",
        "topic": "Dasar Excel untuk Komputasi",
        "agenda": ["Sel, worksheet, workbook", "Referensi relatif dan absolut", "Jenis data dan bilangan", "Error, nol, dan iterasi", "Fungsi Excel dasar"],
        "outcomes": [
            "Membedakan nilai, label, formula, dan format tampilan.",
            "Menggunakan referensi relatif dan absolut dengan benar.",
            "Membaca error Excel sebagai petunjuk debugging.",
            "Membangun tabel hitungan kecil yang dapat diperiksa manual.",
        ],
        "schedule": [
            ["0-10", "Pemanasan: spreadsheet sebagai model hitung", "Satu contoh cell"],
            ["10-28", "Sel, range, worksheet, workbook", "Peta workbook"],
            ["28-45", "Referensi relatif dan absolut", "Salin formula"],
            ["45-60", "Jenis data dan bilangan", "Cek angka vs teks"],
            ["60-75", "Error, nol, dan iterasi", "Lacak #DIV/0!"],
            ["75-90", "Fungsi Excel dasar", "Mini tabel"],
            ["90-100", "Latihan dan exit ticket", "3 jawaban inti"],
        ],
        "sections": [
            ("Excel sebagai Model Hitung", [
                bullet("Cell menyimpan nilai, label, atau formula; format hanya cara menampilkan."),
                bullet("Formula selalu dimulai dengan tanda sama dengan."),
                bullet("Range membantu membaca data sebagai tabel, bukan sel terpisah."),
                bullet("Model yang baik memisahkan input, proses, dan output."),
            ]),
            ("Acu Relatif dan Absolut", [
                bullet("A1 berubah saat formula disalin karena relatif terhadap posisi baru."),
                bullet("$A$1 tetap mengunci kolom dan baris."),
                bullet("A$1 mengunci baris; $A1 mengunci kolom."),
                bullet("Di Excel, tombol cepat untuk toggle referensi adalah F4."),
            ]),
            ("Error dan Nol", [
                bullet("#DIV/0! berarti pembagi kosong atau bernilai nol."),
                bullet("#VALUE! sering muncul karena angka terbaca sebagai teks."),
                bullet("Nol adalah nilai sah; kosong berarti belum ada data."),
                bullet("Debugging dimulai dari cell input, bukan langsung dari output akhir."),
            ]),
            ("Iterasi di Excel", [
                bullet("Iterasi berarti menghitung berulang sampai perubahan kecil atau syarat terpenuhi."),
                bullet("Contoh sipil: coba-coba dimensi sampai kapasitas memenuhi debit rencana."),
                bullet("Batasi iterasi agar proses tidak berjalan tanpa akhir."),
                bullet("Selalu bandingkan hasil iterasi dengan hitungan manual sederhana."),
            ]),
        ],
        "activity": [
            bullet("Buat tabel panjang, lebar, luas untuk tiga baris data."),
            bullet("Salin formula luas dan amati mana referensi yang berubah."),
            bullet("Buat satu kesalahan #DIV/0!, lalu jelaskan penyebabnya."),
        ],
        "exercise": [
            "Apa perbedaan isi cell 10, '10', dan =5+5?",
            "Kapan memakai $B$1 pada formula yang disalin ke banyak baris?",
            "Mengapa nilai nol tidak boleh selalu diperlakukan sama dengan cell kosong?",
        ],
    },
    {
        "file": "03-fungsi-excel-dan-if.pptx",
        "meeting": "Pertemuan 3",
        "topic": "Fungsi Excel dan IF",
        "agenda": ["Fungsi bawaan Excel", "Logika TRUE/FALSE", "IF tunggal", "IF majemuk", "Kasus nilai mahasiswa"],
        "outcomes": [
            "Menulis fungsi Excel dengan argumen yang benar.",
            "Membaca ekspresi logika sebagai TRUE atau FALSE.",
            "Menggunakan IF tunggal dan IF majemuk.",
            "Membuat aturan nilai A, B, C, D, E secara konsisten.",
        ],
        "schedule": [
            ["0-12", "Review formula dan referensi", "Prediksi hasil"],
            ["12-28", "SUM, AVERAGE, MIN, MAX", "Ringkasan data"],
            ["28-43", "Operator logika", "TRUE/FALSE"],
            ["43-60", "IF tunggal", "Lulus/tidak"],
            ["60-78", "IF majemuk", "A-B-C-D-E"],
            ["78-92", "Latihan kasus nilai", "Tabel penilaian"],
            ["92-100", "Latihan dan exit ticket", "3 jawaban inti"],
        ],
        "sections": [
            ("Fungsi Excel", [
                bullet("Fungsi mengubah input menjadi output: =NAMA_FUNGSI(argumen)."),
                bullet("SUM menjumlahkan; AVERAGE mencari rata-rata; MIN/MAX mencari batas."),
                bullet("Argumen dapat berupa angka, cell, atau range."),
                bullet("Gunakan nama fungsi untuk menjelaskan niat hitungan."),
            ]),
            ("Logika TRUE/FALSE", [
                bullet("Ekspresi seperti B2>=60 menghasilkan TRUE atau FALSE."),
                bullet("AND memerlukan semua syarat benar; OR cukup salah satu benar."),
                bullet("Urutan batas nilai harus konsisten agar tidak tumpang tindih."),
                bullet("Logika yang jelas lebih penting daripada formula yang terlihat pendek."),
            ]),
            ("IF Tunggal", [
                bullet("Pola: =IF(syarat, nilai_jika_benar, nilai_jika_salah)."),
                bullet("Contoh: =IF(B2>=60, \"Lulus\", \"Tidak lulus\")."),
                bullet("Selalu uji batas: 59, 60, dan 61."),
                bullet("Kasus batas adalah tempat bug paling sering muncul."),
            ]),
            ("IF Majemuk", [
                bullet("IF bertingkat dipakai saat kategori lebih dari dua."),
                bullet("Contoh nilai: A >= 80, B >= 70, C >= 60, D >= 50, selain itu E."),
                bullet("Mulai dari batas tertinggi agar formula mudah dibaca."),
                bullet("Gunakan tabel bantu jika aturan makin panjang."),
            ]),
        ],
        "activity": [
            bullet("Buat kolom nilai akhir dan huruf mutu untuk 8 mahasiswa."),
            bullet("Uji nilai 49, 50, 59, 60, 69, 70, 79, 80."),
            bullet("Jelaskan satu baris formula IF dengan bahasa manusia."),
        ],
        "exercise": [
            "Prediksi hasil =IF(75>=70,\"B\",\"C\").",
            "Mengapa nilai 80 harus diuji saat membuat kategori A?",
            "Tuliskan aturan IF majemuk untuk A, B, C, D, E dalam kalimat.",
        ],
    },
    {
        "file": "04-algoritma.pptx",
        "meeting": "Pertemuan 4",
        "topic": "Algoritma dan Elemen",
        "agenda": ["Masalah menjadi langkah", "Definisi algoritma", "Input-proses-output", "Flowchart", "Pseudocode dan program"],
        "outcomes": [
            "Memecah persoalan hitungan menjadi langkah berurutan.",
            "Mengenali input, proses, keputusan, pengulangan, dan output.",
            "Menulis pseudocode sederhana sebelum menulis formula atau VBA.",
            "Menghubungkan algoritma dengan program komputer.",
        ],
        "schedule": [
            ["0-10", "Contoh urusan harian dan hitungan sipil", "Langkah manual"],
            ["10-25", "Definisi algoritma", "Ciri algoritma baik"],
            ["25-43", "Input-proses-output", "Sketsa IPO"],
            ["43-62", "Flowchart", "Simbol utama"],
            ["62-80", "Pseudocode", "Versi bahasa manusia"],
            ["80-93", "Terjemah ke Excel/VBA", "Program kecil"],
            ["93-100", "Latihan dan exit ticket", "3 jawaban inti"],
        ],
        "sections": [
            ("Dari Masalah ke Langkah", [
                bullet("Komputer tidak memahami tujuan umum; komputer menjalankan instruksi spesifik."),
                bullet("Tugas kita: mengubah masalah menjadi urutan langkah yang tidak ambigu."),
                bullet("Algoritma boleh ditulis dengan kalimat, flowchart, atau pseudocode."),
                bullet("Sebelum coding, pastikan contoh manualnya benar."),
            ]),
            ("Elemen Algoritma", [
                bullet("Input: data yang diperlukan, termasuk satuan dan batas wajar."),
                bullet("Proses: rumus, transformasi, keputusan, atau pengulangan."),
                bullet("Output: hasil yang akan dipakai pengguna."),
                bullet("Validasi: pembanding untuk mengetahui hasil masuk akal atau tidak."),
            ]),
            ("Flowchart", [
                bullet("Oval: mulai/selesai; jajar genjang: input/output."),
                bullet("Persegi panjang: proses; belah ketupat: keputusan."),
                bullet("Panah menunjukkan urutan eksekusi."),
                bullet("Flowchart membantu melihat cabang dan loop sebelum coding."),
            ]),
            ("Pseudocode", [
                bullet("Pseudocode adalah algoritma dengan bahasa semi-formal."),
                bullet("Contoh: baca panjang, baca lebar, luas = panjang * lebar, tampilkan luas."),
                bullet("Tidak perlu sintaks sempurna, tetapi langkahnya harus bisa dieksekusi."),
                bullet("Pseudocode menjadi jembatan dari ide ke Excel atau VBA."),
            ]),
        ],
        "activity": [
            bullet("Ambil kasus luas penampang persegi panjang."),
            bullet("Tulis input, proses, dan output dalam tiga baris."),
            bullet("Ubah menjadi pseudocode dan flowchart sederhana."),
        ],
        "exercise": [
            "Sebutkan input, proses, dan output untuk hitungan volume beton balok.",
            "Mengapa algoritma perlu diuji dengan angka sederhana sebelum coding?",
            "Kapan flowchart lebih membantu daripada pseudocode?",
        ],
    },
    {
        "file": "05-vba-dan-macro-linier.pptx",
        "meeting": "Pertemuan 5",
        "topic": "VBA dan Macro Linier",
        "agenda": ["Komponen Excel", "Formula di cell", "VBA editor", "Record macro", "Program linier"],
        "outcomes": [
            "Menjelaskan workbook, worksheet, range, macro, module, subroutine.",
            "Merekam macro sederhana dan membaca hasil rekamannya.",
            "Memodifikasi macro rekaman menjadi program linier yang rapi.",
            "Menjalankan langkah hitung Excel dari VBA.",
        ],
        "schedule": [
            ["0-12", "Excel sebagai host VBA", "Peta objek"],
            ["12-28", "Formula cell vs macro", "Bandingkan alur"],
            ["28-45", "VBA editor dan module", "Sub pertama"],
            ["45-62", "Record macro", "Kode hasil rekaman"],
            ["62-80", "Rapikan macro linier", "Macro terbaca"],
            ["80-93", "Latihan hitungan sederhana", "Tombol/run macro"],
            ["93-100", "Latihan dan exit ticket", "3 jawaban inti"],
        ],
        "sections": [
            ("Komponen Excel", [
                bullet("Workbook berisi worksheet; worksheet berisi cell dan range."),
                bullet("Formula cocok untuk model terbuka yang mudah dilihat."),
                bullet("Macro cocok untuk langkah berulang: bersihkan data, isi format, hitung, laporkan."),
                bullet("VBA hidup di dalam aplikasi host seperti Excel atau AutoCAD."),
            ]),
            ("Merekam Macro", [
                bullet("Record macro menangkap operasi manual menjadi kode VBA."),
                bullet("Hasil rekaman sering terlalu panjang, tetapi bagus untuk belajar objek dan perintah."),
                bullet("Baca rekaman sebagai urutan aksi, bukan sebagai kode final."),
                bullet("Setelah paham, rapikan nama, hapus langkah tidak perlu, dan beri struktur."),
            ]),
            ("Program Linier", [
                bullet("Program linier berjalan dari baris pertama ke baris terakhir tanpa cabang."),
                bullet("Pola awal: tulis label, baca input, hitung, tulis output."),
                bullet("Gunakan `Option Explicit` agar salah ketik variabel cepat terlihat."),
                bullet("Uji dengan angka kecil yang bisa dihitung manual."),
            ]),
        ],
        "code": [
            "Option Explicit",
            "",
            "Sub HitungLuasPelat()",
            "    Range(\"A1\").Value = \"Panjang\"",
            "    Range(\"A2\").Value = \"Lebar\"",
            "    Range(\"A3\").Value = \"Luas\"",
            "    Range(\"B3\").Value = Range(\"B1\").Value * Range(\"B2\").Value",
            "End Sub",
        ],
        "activity": [
            bullet("Rekam macro yang memberi header tabel dan warna sederhana."),
            bullet("Buka kode hasil rekaman dan hapus langkah yang tidak perlu."),
            bullet("Tambahkan satu baris hitungan luas dari input B1 dan B2."),
        ],
        "exercise": [
            "Apa bedanya formula di cell dengan macro VBA?",
            "Mengapa hasil record macro biasanya perlu dirapikan?",
            "Apa fungsi `Option Explicit` dalam belajar VBA?",
        ],
    },
    {
        "file": "06-input-output-dan-modularitas.pptx",
        "meeting": "Pertemuan 6",
        "topic": "Input-Output dan Modularitas",
        "agenda": ["Cells dan Range", "Variabel lokal dan public", "Array", "Function", "Subroutine"],
        "outcomes": [
            "Membaca dan menulis data memakai Cells dan Range.",
            "Memakai variabel untuk menyimpan data sementara.",
            "Membedakan Sub dan Function.",
            "Menyusun program kecil yang lebih modular dan mudah diuji.",
        ],
        "schedule": [
            ["0-10", "Review macro linier", "Pola input-output"],
            ["10-28", "Cells, Range, read, write", "Tulis output"],
            ["28-45", "Variabel lokal dan public", "Scope"],
            ["45-60", "Array", "Data berulang"],
            ["60-78", "Function", "Rumus modular"],
            ["78-92", "Subroutine", "Prosedur utama"],
            ["92-100", "Latihan dan exit ticket", "3 jawaban inti"],
        ],
        "sections": [
            ("Input-Output VBA", [
                bullet("`Range(\"B2\").Value` membaca atau menulis satu cell tertentu."),
                bullet("`Cells(baris, kolom)` memudahkan loop karena indeks dapat berubah."),
                bullet("Pisahkan area input dan output agar data lama tidak tertimpa."),
                bullet("Setiap program kecil perlu contoh input dan output acuan."),
            ]),
            ("Variabel dan Scope", [
                bullet("Variabel menyimpan nilai sementara: panjang, lebar, luas, debit."),
                bullet("Variabel lokal hanya berlaku di dalam prosedur tempat ia dibuat."),
                bullet("Variabel public dapat dipakai beberapa prosedur, tetapi lebih sulit dilacak."),
                bullet("Untuk pemula, dahulukan variabel lokal agar debugging lebih jelas."),
            ]),
            ("Array", [
                bullet("Array menyimpan banyak nilai dengan satu nama dan indeks."),
                bullet("Cocok untuk daftar panjang, debit, elevasi, atau nilai mahasiswa."),
                bullet("Indeks membuat program bisa membaca data berulang dalam loop."),
                bullet("Catat batas awal dan akhir indeks agar tidak keluar rentang."),
            ]),
            ("Function dan Subroutine", [
                bullet("Function mengembalikan nilai, misalnya `Luas(P, L)`."),
                bullet("Subroutine menjalankan aksi, misalnya membaca tabel dan menulis laporan."),
                bullet("Modularitas membuat kode mudah diuji bagian per bagian."),
                bullet("Nama fungsi yang jelas adalah dokumentasi pertama."),
            ]),
        ],
        "code": [
            "Option Explicit",
            "",
            "Function LuasPersegiPanjang(p As Double, l As Double) As Double",
            "    LuasPersegiPanjang = p * l",
            "End Function",
            "",
            "Sub TulisLuas()",
            "    Dim p As Double, l As Double",
            "    p = Range(\"B1\").Value",
            "    l = Range(\"B2\").Value",
            "    Range(\"B3\").Value = LuasPersegiPanjang(p, l)",
            "End Sub",
        ],
        "activity": [
            bullet("Buat Function untuk volume balok: panjang * lebar * tinggi."),
            bullet("Buat Sub yang membaca B1:B3 dan menulis hasil di B4."),
            bullet("Uji dengan angka 2, 3, 4 sehingga hasil manualnya 24."),
        ],
        "exercise": [
            "Kapan lebih tepat memakai Cells(i, 2) daripada Range(\"B2\")?",
            "Apa perbedaan utama Function dan Subroutine?",
            "Mengapa variabel public perlu dipakai dengan hati-hati?",
        ],
    },
    {
        "file": "07-userform-dan-control.pptx",
        "meeting": "Pertemuan 7",
        "topic": "UserForm dan Control",
        "agenda": ["UserForm", "Control", "Property", "Event", "Mini aplikasi pra-UTS"],
        "outcomes": [
            "Menjelaskan fungsi UserForm sebagai antarmuka kecil.",
            "Mengenali control umum: Label, TextBox, ComboBox, CommandButton.",
            "Mengatur property penting dan menulis event sederhana.",
            "Menyusun mini aplikasi hitungan yang dapat diverifikasi manual.",
        ],
        "schedule": [
            ["0-10", "Review Function/Sub", "Peta program"],
            ["10-25", "UserForm sebagai antarmuka", "Sketsa form"],
            ["25-42", "Control dan property", "Nama control"],
            ["42-60", "Event", "Klik tombol"],
            ["60-78", "Contoh form hitung luas/volume", "Demo"],
            ["78-93", "Praktik integratif pra-UTS", "Mini aplikasi"],
            ["93-100", "Latihan dan exit ticket", "3 jawaban inti"],
        ],
        "sections": [
            ("UserForm", [
                bullet("UserForm adalah jendela input-output kecil untuk pengguna."),
                bullet("Form berguna saat pengguna tidak perlu melihat semua worksheet."),
                bullet("Form tidak menggantikan validasi; data tetap harus diperiksa."),
                bullet("Desain sederhana lebih baik: sedikit input, tombol jelas, output mudah dibaca."),
            ]),
            ("Control dan Property", [
                bullet("Label memberi keterangan; TextBox menerima input; Button menjalankan perintah."),
                bullet("Property mengatur nama, caption, nilai awal, warna, dan ukuran."),
                bullet("Nama control sebaiknya bermakna: txtPanjang, txtLebar, cmdHitung."),
                bullet("Caption untuk pengguna; Name untuk kode."),
            ]),
            ("Event", [
                bullet("Event adalah kejadian yang memicu kode, misalnya tombol diklik."),
                bullet("Kode event biasanya membaca input, validasi, panggil fungsi, tampilkan output."),
                bullet("Jika input kosong atau bukan angka, tampilkan pesan yang jelas."),
                bullet("Event yang pendek lebih mudah diuji dan dirawat."),
            ]),
            ("Praktik Pra-UTS", [
                bullet("Pilih satu kasus kecil: luas, volume, konversi satuan, atau nilai mahasiswa."),
                bullet("Buat algoritma, worksheet/form, kode, dan hasil pembanding manual."),
                bullet("Verifikasi minimal dengan satu kasus mudah dan satu kasus batas."),
                bullet("Kumpulkan program kecil plus catatan cara mengecek hasilnya."),
            ]),
        ],
        "code": [
            "Private Sub cmdHitung_Click()",
            "    Dim p As Double, l As Double",
            "    p = CDbl(txtPanjang.Value)",
            "    l = CDbl(txtLebar.Value)",
            "    lblHasil.Caption = \"Luas = \" & Format(p * l, \"0.00\")",
            "End Sub",
        ],
        "activity": [
            bullet("Rancang form dengan dua input, satu tombol, dan satu label hasil."),
            bullet("Buat fungsi hitung terpisah dari event tombol."),
            bullet("Uji input normal, input kosong, dan input nol."),
        ],
        "exercise": [
            "Apa bedanya property Name dan Caption pada sebuah Button?",
            "Mengapa kode event sebaiknya tidak terlalu panjang?",
            "Sebutkan dua kasus uji untuk mini aplikasi hitungan luas.",
        ],
    },
]


def build_deck(data):
    src = pick_source(TEMPLATE)
    out_path = os.path.join(OUT_DIR, data["file"])
    tmp_path = out_path + ".tmp"
    shutil.copy(src, tmp_path)
    prs = Presentation(tmp_path)
    delete_all_slides(prs)

    title_slide(
        prs,
        COURSE,
        SUBTITLE,
        data["topic"],
        LECTURER,
        data["meeting"],
        DATE_LABEL,
    )
    agenda_slide(prs, "Materi Pembelajaran", data["agenda"])
    content_slide(prs, "Capaian Pembelajaran", [bullet(x, kind="num", size=15.5) for x in data["outcomes"]], base_size=15.5)
    schedule_slide(prs, data["topic"], data["schedule"])

    divider_slide(prs, 1, "Konsep Inti")
    for title, items in data["sections"]:
        content_slide(prs, title, items, base_size=15.5)

    if data.get("code"):
        divider_slide(prs, 2, "Demo Singkat")
        code_slide(prs, "Contoh Script VBA", data["code"], note="Fokus observasi: input, proses, output, dan lokasi data.")

    content_slide(prs, "Latihan Kelas", data["activity"], base_size=16)
    exercise_slide(prs, data["exercise"])
    content_slide(prs, "Exit Ticket", [
        bullet("Tulis satu konsep yang sudah jelas."),
        bullet("Tulis satu bagian yang masih membingungkan."),
        bullet("Tulis satu contoh penggunaan di teknik sipil yang ingin Anda coba."),
    ], base_size=17)
    closing_slide(prs, "_Mulai dari masalah kecil, buat langkahnya jelas, lalu validasi hasilnya._", "- Algoritma dan Pemrograman Komputer")

    build_and_save(prs, tmp_path, out_path)
    return out_path


def main():
    print("Membangun deck PowerPoint Indonesia...")
    for deck in DECKS:
        build_deck(deck)


if __name__ == "__main__":
    main()
