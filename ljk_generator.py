#!/usr/bin/env python3
"""Generator Lembar Jawab Komputer (LJK) A4 dalam format PDF.

Fitur:
  * Header: judul kuis, nama sekolah, kolom Nama Siswa, Kelas, Date.
  * 4 kotak hitam pekat 1x1 cm di sudut halaman (fiducial markers).
  * QR code di pojok kanan atas berisi metadata: Quiz_ID, Jumlah_Soal, Tanggal.
  * Bulatan jawaban A-E yang menyesuaikan jumlah soal (1 - 100).

Pemakaian cepat:
    python ljk_generator.py --nama-kuis "Kuis Matematika" --jumlah-soal 30 -o ljk_30.pdf
    python ljk_generator.py --demo        # buat contoh 15, 30, 45, 75 soal

Pemakaian dari kode:
    from ljk_generator import generate_ljk_pdf
    generate_ljk_pdf("Kuis Matematika", 30, "ljk_30.pdf")

Dependensi: pip install reportlab "qrcode[pil]"
"""

import argparse
import io
import json
import sys
import uuid
from datetime import date

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas

# --------------------------------------------------------------------------
# Konstanta layout (semua dalam point; mm dari reportlab = 72/25.4 pt)
# --------------------------------------------------------------------------
PAGE_W, PAGE_H = A4  # 595.27 x 841.89 pt (210 x 297 mm)

MARKER_SIZE = 10 * mm      # kotak jangkar 1 x 1 cm
# Jarak kotak jangkar dari tepi kertas. Printer biasa tidak bisa mencetak sampai
# tepi (biasanya terpotong 3-5 mm), jadi 5 mm aman. Ubah ke 0 untuk cetak
# percetakan / full-bleed, sehingga kotak tepat menempel di sudut kertas.
MARKER_MARGIN = 5 * mm

SIDE_MARGIN = 20 * mm      # margin kiri-kanan area konten
OPTIONS = "ABCDE"          # pilihan jawaban
MAX_SOAL = 100             # batas atas jumlah soal (4 kolom x 25 baris)
MAX_ROWS_PER_COL = 25      # baris maksimum per kolom bulatan
BUBBLE_R = 2.3 * mm        # jari-jari bulatan (diameter 4.6 mm)
QR_SIZE = 24 * mm


# --------------------------------------------------------------------------
# Fungsi bantu
# --------------------------------------------------------------------------
def _shorten_to_fit(text, font, size, max_width):
    """Perkecil huruf lalu potong teks (dengan '...') bila terlalu panjang."""
    while size > 9 and stringWidth(text, font, size) > max_width:
        size -= 0.5
    if stringWidth(text, font, size) > max_width:
        while text and stringWidth(text + "...", font, size) > max_width:
            text = text[:-1]
        text += "..."
    return text, size


def _draw_markers(c):
    """Empat kotak hitam pekat 1x1 cm di keempat sudut halaman."""
    c.setFillColorRGB(0, 0, 0)
    xs = (MARKER_MARGIN, PAGE_W - MARKER_MARGIN - MARKER_SIZE)
    ys = (MARKER_MARGIN, PAGE_H - MARKER_MARGIN - MARKER_SIZE)
    for x in xs:
        for y in ys:
            c.rect(x, y, MARKER_SIZE, MARKER_SIZE, stroke=0, fill=1)


def _draw_qr(c, payload, x, y, size):
    """Gambar QR code (library qrcode; cadangan: QR bawaan reportlab)."""
    try:
        import qrcode
    except ImportError:
        print("PERINGATAN: library 'qrcode' tidak ditemukan, memakai QR bawaan "
              "reportlab. Pasang dengan: pip install \"qrcode[pil]\"",
              file=sys.stderr)
        from reportlab.graphics import renderPDF
        from reportlab.graphics.barcode.qr import QrCodeWidget
        from reportlab.graphics.shapes import Drawing

        widget = QrCodeWidget(payload, barLevel="M")
        x1, y1, x2, y2 = widget.getBounds()
        drawing = Drawing(size, size,
                          transform=[size / (x2 - x1), 0, 0, size / (y2 - y1), 0, 0])
        drawing.add(widget)
        renderPDF.draw(drawing, c, x, y)
        return

    qr = qrcode.QRCode(
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=10,
        border=2,  # quiet zone 2 modul sudah termasuk di dalam gambar
    )
    qr.add_data(payload)
    qr.make(fit=True)
    pil_img = qr.make_image(fill_color="black", back_color="white").get_image()
    buf = io.BytesIO()
    pil_img.convert("RGB").save(buf, format="PNG")
    buf.seek(0)
    c.drawImage(ImageReader(buf), x, y, width=size, height=size)


def _draw_header(c, nama_kuis, nama_sekolah, y_bottom_out):
    """Gambar header. Mengembalikan y (point) tempat header berakhir."""
    left = SIDE_MARGIN
    right = PAGE_W - SIDE_MARGIN
    c.setFillColorRGB(0, 0, 0)

    # Judul & sekolah (dibatasi supaya tidak menabrak QR code di kanan)
    text_w = (right - QR_SIZE - 6 * mm) - left
    title, tsize = _shorten_to_fit(nama_kuis, "Helvetica-Bold", 16, text_w)
    c.setFont("Helvetica-Bold", tsize)
    c.drawString(left, PAGE_H - 24 * mm, title)

    school, ssize = _shorten_to_fit(nama_sekolah, "Helvetica", 11, text_w)
    c.setFont("Helvetica", ssize)
    c.drawString(left, PAGE_H - 31 * mm, school)

    c.setFont("Helvetica-Bold", 9)
    c.drawString(left, PAGE_H - 38 * mm, "LEMBAR JAWAB KOMPUTER")

    # Kolom isian siswa
    c.setLineWidth(0.7)
    c.setFont("Helvetica", 10)

    y1 = PAGE_H - 50 * mm
    c.drawString(left, y1, "Nama Siswa")
    c.drawString(left + 22 * mm, y1, ":")
    c.line(left + 25 * mm, y1 - 1.2 * mm, right, y1 - 1.2 * mm)

    y2 = PAGE_H - 59 * mm
    c.drawString(left, y2, "Kelas")
    c.drawString(left + 22 * mm, y2, ":")
    c.line(left + 25 * mm, y2 - 1.2 * mm, left + 75 * mm, y2 - 1.2 * mm)
    c.drawString(left + 85 * mm, y2, "Date")
    c.drawString(left + 97 * mm, y2, ":")
    c.line(left + 100 * mm, y2 - 1.2 * mm, right, y2 - 1.2 * mm)

    # Petunjuk pengisian
    y3 = PAGE_H - 68 * mm
    c.setFont("Helvetica", 8)
    c.drawString(left, y3, "Petunjuk: hitamkan satu bulatan dengan pensil 2B, penuh dan rapi.")
    ex_x = left + 112 * mm
    c.drawString(ex_x, y3, "Benar")
    c.circle(ex_x + 11 * mm, y3 + 0.9 * mm, 1.8 * mm, stroke=1, fill=1)
    c.drawString(ex_x + 18 * mm, y3, "Salah")
    c.circle(ex_x + 29 * mm, y3 + 0.9 * mm, 1.8 * mm, stroke=1, fill=0)
    c.line(ex_x + 27.6 * mm, y3 - 0.5 * mm, ex_x + 30.4 * mm, y3 + 2.3 * mm)

    return PAGE_H - 73 * mm


def _draw_answer_grid(c, jumlah_soal, y_top, y_bottom):
    """Bulatan jawaban: jumlah kolom & tinggi baris menyesuaikan jumlah soal."""
    n_cols = -(-jumlah_soal // MAX_ROWS_PER_COL)   # pembulatan ke atas
    rows = -(-jumlah_soal // n_cols)               # baris per kolom
    header_h = 7 * mm
    label_w = 11 * mm

    pitch_y = min(8.5 * mm, (y_top - header_h - y_bottom) / rows)
    col_w = min((PAGE_W - 2 * SIDE_MARGIN) / n_cols, 60 * mm)
    pitch_x = min(8 * mm, (col_w - label_w - 1 * mm) / len(OPTIONS))
    x_start = (PAGE_W - n_cols * col_w) / 2        # blok kolom di tengah halaman

    c.setStrokeColorRGB(0, 0, 0)
    c.setLineWidth(0.8)

    for col in range(n_cols):
        col_x = x_start + col * col_w

        # Huruf pilihan di atas tiap kolom
        c.setFillColorRGB(0, 0, 0)
        c.setFont("Helvetica-Bold", 9)
        for j, letter in enumerate(OPTIONS):
            cx = col_x + label_w + pitch_x * (j + 0.5)
            c.drawCentredString(cx, y_top - 4.5 * mm, letter)

        for r in range(rows):
            q = col * rows + r + 1                  # nomor soal berurut per kolom
            if q > jumlah_soal:
                break
            cy = y_top - header_h - pitch_y * (r + 0.5)

            # Garis pemisah tipis tiap 5 soal agar mudah dibaca
            if r and r % 5 == 0:
                c.setStrokeGray(0.75)
                c.setLineWidth(0.4)
                sep_y = cy + pitch_y / 2
                c.line(col_x + 2 * mm, sep_y, col_x + col_w - 2 * mm, sep_y)
                c.setStrokeColorRGB(0, 0, 0)
                c.setLineWidth(0.8)

            c.setFillColorRGB(0, 0, 0)
            c.setFont("Helvetica", 9)
            c.drawRightString(col_x + label_w - 3 * mm, cy - 3, str(q))
            for j in range(len(OPTIONS)):
                cx = col_x + label_w + pitch_x * (j + 0.5)
                c.circle(cx, cy, BUBBLE_R, stroke=1, fill=0)


# --------------------------------------------------------------------------
# Fungsi utama
# --------------------------------------------------------------------------
def generate_ljk_pdf(nama_kuis, jumlah_soal, output_filename,
                     nama_sekolah="Nama Sekolah", quiz_id=None, tanggal=None):
    """Buat PDF LJK ukuran A4 dan kembalikan nama file-nya.

    nama_kuis        : judul kuis di header.
    jumlah_soal      : bilangan bulat 1..100 (mis. 15, 30, 45, 75).
    output_filename  : nama file PDF keluaran, mis. "ljk.pdf".
    nama_sekolah     : (opsional) nama sekolah di header.
    quiz_id          : (opsional) ID kuis untuk QR; otomatis dibuat bila kosong.
    tanggal          : (opsional) 'YYYY-MM-DD' atau objek date; default hari ini.
    """
    if (not isinstance(jumlah_soal, int) or isinstance(jumlah_soal, bool)
            or not 1 <= jumlah_soal <= MAX_SOAL):
        raise ValueError(f"jumlah_soal harus bilangan bulat 1 sampai {MAX_SOAL}")

    if tanggal is None:
        tanggal = date.today()
    if isinstance(tanggal, date):
        tanggal = tanggal.isoformat()
    if quiz_id is None:
        quiz_id = f"QZ-{date.today():%Y%m%d}-{uuid.uuid4().hex[:6].upper()}"

    payload = json.dumps(
        {"Quiz_ID": quiz_id, "Jumlah_Soal": jumlah_soal, "Tanggal": tanggal},
        separators=(",", ":"),
    )

    c = canvas.Canvas(output_filename, pagesize=A4)
    c.setTitle(f"LJK - {nama_kuis}")

    _draw_markers(c)
    _draw_qr(c, payload, PAGE_W - SIDE_MARGIN - QR_SIZE, PAGE_H - 18 * mm - QR_SIZE, QR_SIZE)
    y_after_header = _draw_header(c, nama_kuis, nama_sekolah, None)
    _draw_answer_grid(c, jumlah_soal, y_after_header, y_bottom=25 * mm)

    # Footer kecil (di tengah, di antara dua kotak jangkar bawah)
    c.setFillColorRGB(0, 0, 0)
    c.setFont("Helvetica", 7)
    c.drawCentredString(PAGE_W / 2, 17 * mm,
                        f"Quiz ID: {quiz_id}   |   Jumlah Soal: {jumlah_soal}   |   Tanggal: {tanggal}")

    c.showPage()
    c.save()
    return output_filename


# --------------------------------------------------------------------------
# Command line
# --------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(description="Generator PDF Lembar Jawab Komputer (LJK) A4")
    parser.add_argument("--nama-kuis", default="Kuis Matematika")
    parser.add_argument("--nama-sekolah", default="SMA Negeri 1 Contoh")
    parser.add_argument("--jumlah-soal", type=int, default=30)
    parser.add_argument("--quiz-id", default=None)
    parser.add_argument("--tanggal", default=None, help="format YYYY-MM-DD, default hari ini")
    parser.add_argument("-o", "--output", default=None, help="nama file PDF keluaran")
    parser.add_argument("--demo", action="store_true", help="buat contoh 15, 30, 45, 75 soal")
    args = parser.parse_args()

    if args.demo:
        for n in (15, 30, 45, 75):
            out = f"ljk_{n}_soal.pdf"
            generate_ljk_pdf(args.nama_kuis, n, out, nama_sekolah=args.nama_sekolah)
            print("Dibuat:", out)
        return

    out = args.output or f"ljk_{args.jumlah_soal}_soal.pdf"
    generate_ljk_pdf(args.nama_kuis, args.jumlah_soal, out,
                     nama_sekolah=args.nama_sekolah,
                     quiz_id=args.quiz_id, tanggal=args.tanggal)
    print("Dibuat:", out)


if __name__ == "__main__":
    main()
