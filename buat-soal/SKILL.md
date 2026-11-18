# Buat Soal — Skill Guru & Manajemen DOCX

Skill gabungan (all-in-one) untuk membuat dokumen Microsoft Word (.docx), salah satu kumpulan skill untuk guru di repo ini:

1. **Buat Soal Pilihan Ganda** — layout soal dengan opsi A–E dan kunci jawaban
2. **DOCX Manager** — baca, buat, dan manipulasi file DOCX (templating, tabel, formula, gambar, dll)
3. **Wawancara Guru** — interaktif menanyakan kelas, materi, jenis soal, dan jumlah soal

Didesain agar dapat diekspor ke Gemini (atau LLM lain) dengan **satu prompt** — semua dokumentasi dan kode ada di sini, siap salin dan pakai.

## Prerequisites

Library `python-docx` harus terpasang:

```bash
pip install python-docx
```

## Struktur File

- `SKILL.md` — dokumentasi lengkap (file ini)
- `skill_guru.py` — modul Python utama (semua fungsi)
- `requirements.txt` — daftar dependensi

## 1. Buat Soal Pilihan Ganda

Fungsi: `create_soal_docx(output_path, soal_list)`

Layout dokumen yang dihasilkan:

1. **Judul**: "Contoh Soal Pilihan Ganda" (tebal)
2. **Pemisah**: paragraf kosong antara judul dan soal pertama
3. **Soal**: `"{nomor}. {pertanyaan}"` (tidak tebal)
   - Jika ada `statements`, setiap pernyataan ditambahkan dengan nomor `")"` (contoh: `1) teks pernyataan`) untuk membedakan dari nomor soal yang menggunakan `"."`
4. **Opsi**: 5 opsi (A, B, C, D, E), masing-masing pada paragraf terpisah (tidak tebal) — hanya untuk tipe pilihan ganda
5. **Kunci Jawaban**: `"ANS: {huruf}"` (tidak tebal)
6. **Pemisah**: paragraf kosong antara ANS dengan soal berikutnya (tidak ada paragraf kosong setelah soal terakhir)

Format: Font Tahoma 12pt, spacing after 0pt, judul tebal; soal/opsi/ANS tidak tebal.

Parameter `soal_list` adalah list of dict dengan kunci:

- `nomor`: int
- `pertanyaan`: str
- `statements`: list of str (opsional) — banyak pernyataan, diformat dengan nomor `")"`
- `opsi`: list of 5 str `[A, B, C, D, E]`
- `jawaban`: str (`A`/`B`/`C`/`D`/`E`)
- `tipe`: str (opsional) — `'pilihan_ganda_sederhana'` (default), `'pilihan_ganda_komplek'`, `'pilihan_ganda_mcma'`, `'soal_pendek'`, `'banyak_pernyataan'`, `'soal_bergambar'`
- `gambar`: str (opsional) — path file gambar untuk `soal_bergambar`

Contoh penggunaan:

```python
from skill_guru import create_soal_docx

soal_list = [
    {
        'nomor': 1,
        'pertanyaan': 'Ibukota negara Indonesia adalah',
        'opsi': ['Bandung', 'Surabaya', 'Jakarta', 'Medan', 'Yogyakarta'],
        'jawaban': 'C',
    },
    {
        'nomor': 2,
        'pertanyaan': 'Pernyataan tentang Pancasila yang benar adalah',
        'statements': [
            'Ketuhanan Yang Maha Esa adalah silakan ke-1',
            'Kemanusiaan yang Adil dan Beradab adalah silakan ke-2',
            'Persatuan Indonesia adalah silakan ke-3',
        ],
        'opsi': ['1 dan 2 benar', '1 dan 3 benar', '2 dan 3 benar', '1, 2, dan 3 benar', 'Hanya 1 benar'],
        'jawaban': 'D',
    },
]

create_soal_docx('soal.docx', soal_list)
```

## 2. Wawancara Guru

Fungsi: `wawancara_guru()`

Interaktif menanyakan informasi dari guru untuk menyiapkan pembuatan soal. Mengembalikan dict dengan kunci:

- `kelas`: str — kelas / jenjang
- `materi`: str — materi atau tujuan pembelajaran
- `tujuan_pembelajaran`: list of str — daftar tujuan pembelajaran (default dicari oleh AI)
- `jenis_soal`: str — jenis soal yang dipilih
- `jumlah_pendek`: int — jumlah soal pendek (default 50)
- `jumlah_stimulus`: int — jumlah soal berstimulus (default 20)
- `jumlah_banyak_pernyataan`: int — jumlah banyak pernyataan (default 20)
- `jumlah_gambar`: int — jumlah soal bergambar/grafik (default 10)

Alur wawancara:

1. **Kelas / Jenjang** — guru menuliskan kelas atau jenjang (contoh: Kelas 7 SMP)
2. **Materi / Tujuan pembelajaran** — guru menuliskan materi; opsi default dicari oleh AI atau ditulis sendiri
3. **Jenis soal** — pilihan ganda sederhana (default, 5 opsi), pilihan ganda komplek, atau pilihan ganda MCMA
4. **Jumlah soal** — default: 50 pendek, 20 stimulus, 20 banyak pernyataan, 10 gambar/grafik

Contoh penggunaan:

```python
from skill_guru import wawancara_guru

info = wawancara_guru()
print(info)
# Output contoh:
# {
#     'kelas': 'Kelas 7 SMP',
#     'materi': 'Sistem Pencernaan',
#     'tujuan_pembelajaran': ['(dicari oleh AI berdasarkan materi)'],
#     'jenis_soal': 'pilihan_ganda_sederhana',
#     'jumlah_pendek': 50,
#     'jumlah_stimulus': 20,
#     'jumlah_banyak_pernyataan': 20,
#     'jumlah_gambar': 10,
# }
```

## 3. Alur Output (Penting)

Saat pengguna meminta pembuatan soal pilihan ganda, assistant **harus** menghasilkan file `.docx` secara otomatis, bukan hanya teks dalam chat. Alur yang dianjurkan:

### 3.1 Default: buat file `.docx` otomatis

Setiap kali pengguna meminta soal (misal: "buatkan 5 soal tentang sistem pencernaan"), assistant harus:

1. Susun `soal_list` (list of dict) sesuai format di bagian 1.
2. Panggil `create_soal_docx('nama_file.docx', soal_list)` menggunakan `skill_guru.py`.
3. Kirimkan file `.docx` hasil generate kepada pengguna sebagai attachment.
4. Sertakan ringkasan singkat isi soal dalam chat (opsional).

### 3.2 Alternatif: tanya format output

Jika tidak yakin apakah pengguna ingin file `.docx` atau hanya teks dalam chat, assistant boleh menanyakan terlebih dahulu:

> "Apakah Anda ingin soal dalam format file `.docx` (dapat diunduh) atau hanya teks dalam chat?"

Tunggu balasan pengguna sebelum melanjutkan.

### 3.3 Contoh pemanggilan otomatis

```python
from skill_guru import create_soal_docx

soal_list = [
    {
        'nomor': 1,
        'pertanyaan': 'Ibukota negara Indonesia adalah',
        'opsi': ['Bandung', 'Surabaya', 'Jakarta', 'Medan', 'Yogyakarta'],
        'jawaban': 'C',
    },
    # ... tambahkan soal lain
]

create_soal_docx('soal_sistem_pencernaan.docx', soal_list)
# File siap dikirim sebagai attachment ke pengguna
```

## 4. DOCX Manager

Semua fungsi tersedia di `skill_guru.py` dan dapat dipanggil langsung dari Python.

### 4.1 Baca DOCX

`read_docx(filepath)` — cetak seluruh teks paragraf dari file DOCX.

### 4.2 Buat DOCX sederhana

`create_docx(filepath, text)` — buat file DOCX baru dari teks (baris dipisahkan dengan `\n`).

### 4.3 Tambah konten ke DOCX

`add_to_docx(filepath, text)` — tambahkan paragraf ke file DOCX yang sudah ada.

### 4.4 Buat DOCX dengan format

`create_formatted_docx(filepath, content_json)` — buat dokumen dari JSON dengan kontrol format granular (bold, italic, underline, style).

Format `content_json`:

```json
{
  "paragraphs": [
    {"text": "Judul Dokumen", "style": "Heading 1"},
    {"text": "Isi paragraf biasa", "style": "Normal"},
    {"text": "Teks tebal", "bold": true},
    {"text": "Teks miring", "italic": true}
  ]
}
```

### 4.5 Buat Tabel

`create_table(filepath, rows, cols, data=None, caption=None)` — buat tabel grid dengan header berwarna biru tua dan teks putih tebal.

### 4.6 Templating (Placeholder Replacement)

`create_template(filepath, template_text, replacements)` — buat dokumen dari template dengan mengganti placeholder `{key}`.

Contoh:

```python
create_template('surat.docx', 'Halo {nama}, acara pada {tanggal}.', {'nama': 'Andi', 'tanggal': 'Besok'})
```

### 4.7 Insert Formula (Rumus)

`insert_formula(filepath, formula, position='end')` — sisipkan rumus matematika ke dalam DOCX (menggunakan OMML/EQ field code).

### 4.8 Insert Gambar

`insert_image(filepath, image_path, width=None, position='end')` — sisipkan gambar ke dalam dokumen yang sudah ada.

### 4.9 Dokumen Lengkap (Comprehensive)

`create_comprehensive_docx(output_path)` — contoh dokumen lengkap dengan judul, paragraf Lorem Ipsum, tabel data, rumus, dan gambar.

## Contoh Lengkap (Satu File)

Lihat `skill_guru.py` untuk implementasi lengkap semua fungsi di atas. Semua fungsi independen dan hanya bergantung pada `python-docx`.

## Export ke Gemini (Satu Prompt)

Salin seluruh isi `SKILL.md` dan `skill_guru.py` ke dalam satu prompt Gemini. Skill ini siap digunakan — tidak perlu dependensi eksternal lain selain `python-docx`.

## Lisensi

Lihat file `LICENSE` di root repository.
