# skill-guru

Kumpulan skill yang berguna untuk guru, siap diekspor ke Gemini (atau LLM lain) dengan satu prompt.

## Skill yang tersedia

| Folder | Deskripsi |
|--------|-----------|
| `buat-soal/` | Membuat soal pilihan ganda DOCX dan mengelola file DOCX (templating, tabel, formula, gambar) |

## Cara penggunaan

Setiap skill adalah folder mandiri yang berisi:
- `SKILL.md` — dokumentasi lengkap
- `skill_guru.py` — modul Python utama
- `requirements.txt` — dependensi (`python-docx`)

Untuk mengekspor ke Gemini, salin isi `SKILL.md` dan `skill_guru.py` ke dalam satu prompt.

## Pemasangan dependensi

```bash
pip install -r buat-soal/requirements.txt
```

## Lisensi

Lihat file `LICENSE` di root repository.
