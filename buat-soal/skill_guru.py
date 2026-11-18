#!/usr/bin/env python3
"""
Skill Guru — Buat Soal & Manajemen DOCX (all-in-one)

Modul ini menggabungkan:
1. create_soal_docx  — buat dokumen .docx berisi soal pilihan ganda (layout disesuaikan)
2. DOCX Manager      — baca, buat, dan manipulasi file DOCX (templating, tabel, formula, gambar)
3. create_comprehensive_docx — contoh dokumen lengkap (paragraf, tabel, formula, gambar)
4. Wawancara Guru    — interaktif menanyakan kelas, materi, jenis soal, jumlah soal

Semua fungsi hanya bergantung pada python-docx.
Didesain untuk diekspor ke Gemini (atau LLM lain) dengan satu prompt.
"""

import sys
import json
import os
import re

try:
    from docx import Document
    from docx.shared import Pt, Inches, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
except ImportError:
    print("ERROR: python-docx belum terinstall. Jalankan: pip install python-docx")
    sys.exit(1)


# ---------------------------------------------------------------------------
# Warna standar header tabel (biru tua dengan teks putih)
# ---------------------------------------------------------------------------
HEADER_BG_COLOR = "2F5496"
HEADER_TEXT_COLOR = RGBColor(0xFF, 0xFF, 0xFF)


# ===========================================================================
# Utilitas umum
# ===========================================================================

def set_cell_shading(cell, fill_hex):
    """Set warna background (shading) pada sebuah sel tabel."""
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn('w:shd'))
    if shd is None:
        shd = OxmlElement('w:shd')
        tc_pr.append(shd)
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_hex)


def color_table_header(table, fill_hex=HEADER_BG_COLOR, text_color=HEADER_TEXT_COLOR):
    """Beri warna pada baris header tabel dan buat teks tebal."""
    if not table.rows:
        return
    header_row = table.rows[0]
    for cell in header_row.cells:
        set_cell_shading(cell, fill_hex)
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.bold = True
                run.font.color.rgb = text_color


def ensure_dir(filepath):
    """Buat direktori induk jika belum ada."""
    directory = os.path.dirname(filepath)
    if directory:
        os.makedirs(directory, exist_ok=True)


# ===========================================================================
# 1. BUAT SOAL PILIHAN GANDA
# ===========================================================================

def set_paragraph_format(paragraph):
    """Set spacing after 0pt dan before 0pt untuk paragraf."""
    paragraph_format = paragraph.paragraph_format
    paragraph_format.space_after = Pt(0)
    paragraph_format.space_before = Pt(0)


def set_run_format(run, bold=False):
    """Set font Tahoma ukuran 12 untuk run."""
    run.font.name = 'Tahoma'
    run.font.size = Pt(12)
    run.bold = bold


def create_soal_docx(output_path, soal_list):
    """
    Buat dokumen .docx dengan daftar soal pilihan ganda 5 opsi.

    soal_list: list of dict dengan kunci:
        - nomor: int
        - pertanyaan: str
        - statements: list of str (opsional) — banyak pernyataan, akan diformat
          dengan nomner ")" (contoh: "1) teks pernyataan")
        - opsi: list of 5 str [A, B, C, D, E]
        - jawaban: str (A/B/C/D/E)
        - tipe: str (opsional) — 'pilihan_ganda_sederhana' (default), 'pilihan_ganda_komplek',
          'pilihan_ganda_mcma', 'soal_pendek', 'banyak_pernyataan', 'soal_bergambar'
        - gambar: str (opsional) — path file gambar untuk soal_bergambar
    """
    doc = Document()

    # Judul (tebal)
    p_judul = doc.add_paragraph()
    set_paragraph_format(p_judul)
    run_judul = p_judul.add_run('Contoh Soal Pilihan Ganda')
    set_run_format(run_judul, bold=True)

    # Paragraf kosong pemisah antara judul dan soal pertama
    doc.add_paragraph()

    for i, s in enumerate(soal_list):
        # Soal (tidak tebal)
        p_soal = doc.add_paragraph()
        set_paragraph_format(p_soal)

        # Bangun teks soal
        soal_teks = f"{s['nomor']}. {s['pertanyaan']}"

        # Tambahkan statements jika ada (dengan nomner ")")
        statements = s.get('statements')
        if statements:
            for idx, stmt in enumerate(statements, start=1):
                soal_teks += f"\n{idx}) {stmt}"
            soal_teks += " ..."

        run_soal = p_soal.add_run(soal_teks)
        set_run_format(run_soal, bold=False)

        # Opsi A-E (tidak tebal) — hanya untuk tipe pilihan ganda
        tipe = s.get('tipe', 'pilihan_ganda_sederhana')
        if tipe in ('pilihan_ganda_sederhana', 'pilihan_ganda_komplek', 'pilihan_ganda_mcma'):
            huruf_list = ['A', 'B', 'C', 'D', 'E']
            for huruf, opsi_teks in zip(huruf_list, s['opsi']):
                p_opsi = doc.add_paragraph(f'{huruf}. {opsi_teks}')
                set_paragraph_format(p_opsi)
                run_opsi = p_opsi.runs[0]
                set_run_format(run_opsi, bold=False)

        # Kunci jawaban (tidak tebal)
        p_ans = doc.add_paragraph()
        set_paragraph_format(p_ans)
        run_ans = p_ans.add_run(f'ANS: {s["jawaban"]}')
        set_run_format(run_ans, bold=False)

        # Paragraf kosong pemisah antara ANS dengan soal berikutnya
        if i < len(soal_list) - 1:
            doc.add_paragraph()

    ensure_dir(output_path)
    doc.save(output_path)
    print(f'Dokumen soal berhasil dibuat: {output_path}')


def wawancara_guru():
    """
    Interaktif menanyakan informasi dari guru untuk menyiapkan pembuatan soal.

    Mengembalikan dict dengan kunci:
        - kelas: str
        - materi: str
        - tujuan_pembelajaran: list of str
        - jenis_soal: str
        - jumlah_pendek: int
        - jumlah_stimulus: int
        - jumlah_banyak_pernyataan: int
        - jumlah_gambar: int
    """
    print('=== Wawancara Guru — Persiapan Pembuatan Soal ===')

    kelas = input('Kelas / Jenjang (contoh: Kelas 7 SMP): ').strip()
    materi = input('Materi atau tujuan pembelajaran: ').strip()

    print('\nTujuan pembelajaran:')
    print('  1. Saya (AI) akan mencari default berdasarkan materi')
    print('  2. Saya akan menuliskan sendiri')
    pilihan_tujuan = input('Pilih (1/2): ').strip()

    if pilihan_tujuan == '2':
        tujuan_input = input('Tuliskan semua tujuan pembelajaran (pisahkan dengan titik koma ;): ').strip()
        tujuan_pembelajaran = [t.strip() for t in tujuan_input.split(';') if t.strip()]
    else:
        tujuan_pembelajaran = ['(dicari oleh AI berdasarkan materi)']

    print('\nJenis soal:')
    print('  1. Pilihan ganda sederhana (default, 5 opsi)')
    print('  2. Pilihan ganda komplek')
    print('  3. Pilihan ganda MCMA (multiple correct answers)')
    pilihan_jenis = input('Pilih (1/2/3): ').strip()

    jenis_map = {
        '1': 'pilihan_ganda_sederhana',
        '2': 'pilihan_ganda_komplek',
        '3': 'pilihan_ganda_mcma',
    }
    jenis_soal = jenis_map.get(pilihan_jenis, 'pilihan_ganda_sederhana')

    print('\nJumlah soal (default: 50 pendek, 20 stimulus, 20 banyak pernyataan, 10 gambar):')
    jumlah_pendek = input('Soal pendek [50]: ').strip() or '50'
    jumlah_stimulus = input('Soal berstimulus [20]: ').strip() or '20'
    jumlah_banyak_pernyataan = input('Banyak pernyataan [20]: ').strip() or '20'
    jumlah_gambar = input('Soal bergambar/grafik [10]: ').strip() or '10'

    return {
        'kelas': kelas,
        'materi': materi,
        'tujuan_pembelajaran': tujuan_pembelajaran,
        'jenis_soal': jenis_soal,
        'jumlah_pendek': int(jumlah_pendek),
        'jumlah_stimulus': int(jumlah_stimulus),
        'jumlah_banyak_pernyataan': int(jumlah_banyak_pernyataan),
        'jumlah_gambar': int(jumlah_gambar),
    }


# ===========================================================================
# 2. DOCX MANAGER
# ===========================================================================

def read_docx(filepath):
    """Baca seluruh teks dari file DOCX."""
    if not os.path.exists(filepath):
        print(f"ERROR: File tidak ditemukan: {filepath}")
        return False
    try:
        doc = Document(filepath)
        for paragraph in doc.paragraphs:
            print(paragraph.text)
        return True
    except Exception as e:
        print(f"ERROR: Gagal membaca file: {e}")
        return False


def create_docx(filepath, text):
    """Buat file DOCX baru dengan teks."""
    try:
        doc = Document()
        lines = text.replace('\\n', '\n').split('\n')
        for line in lines:
            if line.strip():
                doc.add_paragraph(line)
        ensure_dir(filepath)
        doc.save(filepath)
        print(f"SUCCESS: File dibuat: {filepath}")
        return True
    except Exception as e:
        print(f"ERROR: Gagal membuat file: {e}")
        return False


def add_to_docx(filepath, text):
    """Tambahkan teks ke file DOCX yang sudah ada."""
    if not os.path.exists(filepath):
        print(f"ERROR: File tidak ditemukan: {filepath}")
        return False
    try:
        doc = Document(filepath)
        lines = text.replace('\\n', '\n').split('\n')
        for line in lines:
            if line.strip():
                doc.add_paragraph(line)
        doc.save(filepath)
        print(f"SUCCESS: Konten ditambahkan ke: {filepath}")
        return True
    except Exception as e:
        print(f"ERROR: Gagal menambah konten: {e}")
        return False


def create_formatted_docx(filepath, content_json):
    """Buat DOCX dengan format dari JSON."""
    try:
        data = json.loads(content_json)
        doc = Document()

        for item in data.get('paragraphs', []):
            text = item.get('text', '')
            style = item.get('style', 'Normal')
            bold = item.get('bold', False)
            italic = item.get('italic', False)
            underline = item.get('underline', False)

            p = doc.add_paragraph(
                style=style if style in ['Heading 1', 'Heading 2', 'Heading 3', 'Normal'] else 'Normal'
            )
            run = p.add_run(text)
            run.bold = bold
            run.italic = italic
            run.underline = underline

        ensure_dir(filepath)
        doc.save(filepath)
        print(f"SUCCESS: File dengan format dibuat: {filepath}")
        return True
    except json.JSONDecodeError as e:
        print(f"ERROR: JSON tidak valid: {e}")
        return False
    except Exception as e:
        print(f"ERROR: Gagal membuat file: {e}")
        return False


def create_table(filepath, rows, cols, data=None, caption=None):
    """
    Buat tabel dalam DOCX.

    Args:
        filepath: Path file output
        rows: Jumlah baris
        cols: Jumlah kolom
        data: List of lists dengan isi tabel (opsional)
        caption: Caption tabel (opsional)
    """
    try:
        doc = Document()

        if caption:
            doc.add_paragraph(caption, style='Heading 2')

        table = doc.add_table(rows=rows, cols=cols)
        table.style = 'Table Grid'

        if data:
            for i, row_data in enumerate(data):
                for j, cell_data in enumerate(row_data):
                    if j < len(table.rows[i].cells):
                        table.rows[i].cells[j].text = str(cell_data)

        color_table_header(table)

        ensure_dir(filepath)
        doc.save(filepath)
        print(f"SUCCESS: Tabel dibuat: {filepath}")
        return True
    except Exception as e:
        print(f"ERROR: Gagal membuat tabel: {e}")
        return False


def insert_formula(filepath, formula, position='end'):
    """
    Insert formula matematika ke dalam DOCX.

    Args:
        filepath: Path file DOCX
        formula: Rumus matematika (misal: "=SUM(A1:B2)", "=A1+B2")
        position: 'end' atau 'start'
    """
    if not os.path.exists(filepath):
        print(f"ERROR: File tidak ditemukan: {filepath}")
        return False
    try:
        doc = Document(filepath)
        p = doc.add_paragraph()
        p.add_run().add_field_code(f'EQ \\* MERGEFORMAT {formula}')
        doc.save(filepath)
        print(f"SUCCESS: Formula ditambahkan ke: {filepath}")
        return True
    except Exception as e:
        print(f"ERROR: Gagal menambah formula: {e}")
        return False


def insert_image(filepath, image_path, width=None, position='end'):
    """
    Insert gambar ke dalam DOCX.

    Args:
        filepath: Path file DOCX
        image_path: Path file gambar
        width: Lebar gambar dalam inch (opsional)
        position: 'end' atau 'start'
    """
    if not os.path.exists(filepath):
        print(f"ERROR: File tidak ditemukan: {filepath}")
        return False
    if not os.path.exists(image_path):
        print(f"ERROR: Gambar tidak ditemukan: {image_path}")
        return False
    try:
        doc = Document(filepath)
        if width:
            doc.add_picture(image_path, width=Inches(width))
        else:
            doc.add_picture(image_path)
        doc.save(filepath)
        print(f"SUCCESS: Gambar ditambahkan ke: {filepath}")
        return True
    except Exception as e:
        print(f"ERROR: Gagal menambah gambar: {e}")
        return False


def create_template(filepath, template_text, replacements):
    """
    Buat dokumen dari template dengan replacement placeholder.

    Args:
        filepath: Path file output
        template_text: Template dengan placeholder {placeholder}
        replacements: Dictionary dengan key sebagai placeholder dan value sebagai pengganti
    """
    try:
        doc = Document()
        content = template_text
        for key, value in replacements.items():
            content = content.replace(f'{{{key}}}', str(value))

        lines = content.replace('\\n', '\n').split('\n')
        for line in lines:
            if line.strip():
                doc.add_paragraph(line)

        ensure_dir(filepath)
        doc.save(filepath)
        print(f"SUCCESS: Template dibuat: {filepath}")
        return True
    except Exception as e:
        print(f"ERROR: Gagal membuat template: {e}")
        return False


# ===========================================================================
# 3. DOKUMEN LENGKAP (COMPREHENSIVE)
# ===========================================================================

def create_comprehensive_docx(output_path):
    """Buat file DOCX lengkap dengan semua elemen (contoh demonstrasi)."""
    doc = Document()

    # 1. Judul
    title = doc.add_paragraph('Dokumen Lengkap dengan Tabel dan Gambar', style='Heading 1')
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 2. Paragraf Lorem Ipsum
    lorem = """
Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.

Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.

Sed ut perspiciatis unde omnis iste natus error sit voluptatem accusantium doloremque laudantium, totam rem aperiam, eaque ipsa quae ab illo inventore veritatis et quasi architecto beatae vitae dicta sunt explicabo.

Nemo enim ipsam voluptatem quia voluptas sit aspernatur aut odit aut fugit, sed quia consequuntur magni dolores eos qui ratione voluptatem sequi nesciunt.
"""
    doc.add_paragraph(lorem.strip())

    # 3. Bagian Tabel
    doc.add_paragraph('Tabel Data Penjualan', style='Heading 2')

    table = doc.add_table(rows=5, cols=4)
    table.style = 'Table Grid'

    headers = ['No', 'Produk', 'Jumlah', 'Harga Satuan']
    for i, header in enumerate(headers):
        table.rows[0].cells[i].text = header

    data = [
        ['1', 'Laptop', '10', 'Rp 15.000.000'],
        ['2', 'Mouse', '50', 'Rp 150.000'],
        ['3', 'Keyboard', '30', 'Rp 300.000'],
        ['4', 'Monitor', '15', 'Rp 2.500.000'],
    ]

    for i, row_data in enumerate(data, start=1):
        for j, cell_data in enumerate(row_data):
            table.rows[i].cells[j].text = cell_data

    color_table_header(table)

    doc.add_paragraph('Total Penjualan: =SUM(D2:D5)', style='Normal')

    # 4. Bagian Gambar
    doc.add_paragraph('Gambar Produk', style='Heading 2')

    try:
        from PIL import Image, ImageDraw, ImageFont
        img = Image.new('RGB', (400, 300), color='lightblue')
        d = ImageDraw.Draw(img)
        try:
            font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 24)
        except Exception:
            font = ImageFont.load_default()
        d.text((100, 130), "Produk Laptop", fill='black', font=font)
        d.text((100, 160), "Rp 15.000.000", fill='blue', font=font)
        temp_img_path = '/tmp/temp_product.png'
        img.save(temp_img_path)
        doc.add_picture(temp_img_path, width=Inches(4))
        os.remove(temp_img_path)
    except ImportError:
        doc.add_paragraph('[Gambar Laptop - Rp 15.000.000]', style='Normal')
        doc.add_paragraph('Catatan: Gambar tidak tersedia karena library PIL tidak terinstall')

    # 5. Ringkasan
    doc.add_paragraph('Ringkasan', style='Heading 2')
    summary = """
Dokumen ini menunjukkan kemampuan Skill Guru untuk:
1. Menambahkan paragraf teks dengan Lorem Ipsum
2. Membuat tabel dengan format grid
3. Menyisipkan rumus matematika
4. Menambahkan gambar ke dalam dokumen

Semua fitur ini dapat digabungkan untuk membuat dokumen yang lengkap dan profesional.
"""
    doc.add_paragraph(summary.strip())

    ensure_dir(output_path)
    doc.save(output_path)
    print(f"SUCCESS: Dokumen lengkap dibuat: {output_path}")
    return True


# ===========================================================================
# CLI entry point
# ===========================================================================

def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    command = sys.argv[1]

    if command == 'create_soal_docx' and len(sys.argv) >= 4:
        # create_soal_docx <output_path> <soal_json>
        soal_list = json.loads(sys.argv[3])
        create_soal_docx(sys.argv[2], soal_list)
    elif command == 'read_docx' and len(sys.argv) >= 3:
        read_docx(sys.argv[2])
    elif command == 'create_docx' and len(sys.argv) >= 4:
        create_docx(sys.argv[2], sys.argv[3])
    elif command == 'add_to_docx' and len(sys.argv) >= 4:
        add_to_docx(sys.argv[2], sys.argv[3])
    elif command == 'create_formatted_docx' and len(sys.argv) >= 4:
        create_formatted_docx(sys.argv[2], sys.argv[3])
    elif command == 'create_table' and len(sys.argv) >= 5:
        filepath = sys.argv[2]
        rows = int(sys.argv[3])
        cols = int(sys.argv[4])
        data = json.loads(sys.argv[5]) if len(sys.argv) > 5 else None
        caption = sys.argv[6] if len(sys.argv) > 6 else None
        create_table(filepath, rows, cols, data, caption)
    elif command == 'insert_formula' and len(sys.argv) >= 4:
        insert_formula(sys.argv[2], sys.argv[3])
    elif command == 'insert_image' and len(sys.argv) >= 4:
        width = float(sys.argv[4]) if len(sys.argv) > 4 else None
        insert_image(sys.argv[2], sys.argv[3], width=width)
    elif command == 'create_template' and len(sys.argv) >= 4:
        create_template(sys.argv[2], sys.argv[3], json.loads(sys.argv[4]))
    elif command == 'create_comprehensive_docx' and len(sys.argv) >= 3:
        create_comprehensive_docx(sys.argv[2])
    else:
        print(f"ERROR: Command tidak dikenal atau argumen tidak lengkap: {command}")
        print(__doc__)
        sys.exit(1)


if __name__ == '__main__':
    main()
