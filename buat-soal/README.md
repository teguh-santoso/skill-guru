# Buat Soal

Skill untuk membuat soal pilihan ganda DOCX dan mengelola file DOCX (templating, tabel, formula, gambar). Bagian dari kumpulan skill untuk guru.

## File

- `SKILL.md` — dokumentasi lengkap
- `skill_guru.py` — modul Python utama
- `requirements.txt` — dependensi (`python-docx`)

## Pemasangan

```bash
pip install -r requirements.txt
```

## Penggunaan

```python
from skill_guru import create_soal_docx, read_docx, create_docx, create_table, insert_image, insert_formula, create_template
```

Lihat `SKILL.md` untuk dokumentasi lengkap dan contoh penggunaan.

## Export ke Gemini

Salin `SKILL.md` + `skill_guru.py` ke satu prompt Gemini.
