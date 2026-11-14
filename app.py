#!/usr/bin/env python3
"""Pemeriksa LJK OMR - Streamlit, mobile-first (gaya ZipGrade), siap Streamlit Community Cloud.

Menu bawah (bottom navigation):  🏠 Kuis  ·  📸 Scan  ·  👥 Siswa  ·  📊 Rekap

Modul yang dipakai (harus satu folder dengan app.py):
    ljk_generator.py    -> membuat PDF LJK
    database_manager.py -> database SQLite + impor/ekspor Excel + cadangan
    omr_engine.py       -> membaca foto LJK (OpenCV) dan menilainya

Jalankan lokal :  streamlit run app.py
Deploy         :  lihat PANDUAN_DEPLOY.md

Kata sandi (opsional, disarankan untuk aplikasi publik): isi `APP_PASSWORD` di
Settings > Secrets pada Streamlit Cloud, atau sebagai environment variable.
"""

import base64
import hashlib
import hmac
import html
import io
import json
import os
import re
import tempfile
import warnings
from datetime import date, datetime

import cv2
import numpy as np
import pandas as pd
import streamlit as st

import database_manager as dm
import ljk_generator as gen
import omr_engine as omr

# --------------------------------------------------------------------------
# Konstanta
# --------------------------------------------------------------------------
MIN_SOAL, MAX_SOAL = 1, 100
CHOICES = list(dm.VALID_OPTIONS)  # A-E, sama dengan tata letak LJK

MODE_100 = "Sama rata, skor maksimum 100"
MODE_1 = "Sama rata, 1 poin per soal"
MODE_MANUAL = "Atur manual per soal"
WEIGHT_MODES = [MODE_100, MODE_1, MODE_MANUAL]

ALL_CLASSES = "(Semua kelas)"
DEFAULT_FILL_THRESHOLD = 0.40
LOW_RES_SIDE = 720          # sisi terpendek foto (px) di bawah ini = peringatan resolusi rendah
MAX_IMAGE_SIDE = 3200       # foto lebih besar diperkecil agar hemat memori di server

NAV_KUIS = "🏠 Kuis"
NAV_SCAN = "📸 Scan"
NAV_SISWA = "👥 Siswa"
NAV_REKAP = "📊 Rekap"
NAV_ITEMS = [NAV_KUIS, NAV_SCAN, NAV_SISWA, NAV_REKAP]

SRC_CAMERA = "📷 Kamera"
SRC_UPLOAD = "📁 Unggah foto"


# ==========================================================================
# LAPISAN LOGIKA (tanpa tampilan streamlit, sehingga bisa diuji terpisah)
# ==========================================================================
def _rows(sql, params=()):
    with dm.get_connection() as conn:
        return [dict(r) for r in conn.execute(sql, params).fetchall()]


def _safe_name(text, fallback="x"):
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "_", str(text)).strip("_")
    return cleaned or fallback


def get_counts():
    return _rows(
        "SELECT (SELECT COUNT(*) FROM classes)      AS kelas,"
        "       (SELECT COUNT(*) FROM students)     AS siswa,"
        "       (SELECT COUNT(*) FROM quizzes)      AS kuis,"
        "       (SELECT COUNT(*) FROM scan_results) AS scan"
    )[0]


def list_quizzes():
    return _rows(
        "SELECT q.id, q.quiz_name AS name, q.total_questions AS total_questions,"
        "       (SELECT COUNT(DISTINCT r.student_id) FROM scan_results r"
        "        WHERE r.quiz_id = q.id) AS peserta "
        "FROM quizzes q ORDER BY q.id DESC"
    )


def get_quiz(quiz_id):
    """Data kuis lengkap: kunci jawaban & bobot sebagai list (indeks 0 = soal 1)."""
    rows = _rows("SELECT * FROM quizzes WHERE id = ?", (quiz_id,))
    if not rows:
        raise ValueError(f"Kuis dengan id {quiz_id!r} tidak ditemukan")
    row = rows[0]
    n = row["total_questions"]
    key_map = json.loads(row["answer_key_json"])
    weight_map = json.loads(row["weights_json"]) if row["weights_json"] else {}
    keys = [key_map[str(i)] for i in range(1, n + 1)]
    weights = [float(weight_map.get(str(i), 1.0)) for i in range(1, n + 1)]
    return {
        "id": row["id"],
        "name": row["quiz_name"],
        "total_questions": n,
        "key": keys,
        "weights": weights,
        "max_score": sum(weights),
    }


def list_students():
    rows = _rows(
        "SELECT s.student_id AS NIS, s.full_name AS Nama, c.class_name AS Kelas "
        "FROM students s JOIN classes c ON c.id = s.class_id "
        "ORDER BY c.class_name COLLATE NOCASE, s.full_name COLLATE NOCASE"
    )
    return pd.DataFrame(rows, columns=["NIS", "Nama", "Kelas"])


def scanned_nis(quiz_id):
    rows = _rows(
        "SELECT DISTINCT s.student_id AS nis FROM scan_results r "
        "JOIN students s ON s.id = r.student_id WHERE r.quiz_id = ?",
        (quiz_id,),
    )
    return {r["nis"] for r in rows}


def next_unscanned(order, scanned, after=None):
    """NIS berikutnya (urutan daftar, memutar ke awal) yang belum di-scan.

    order   : daftar NIS sesuai urutan tampil
    scanned : himpunan NIS yang sudah punya nilai
    after   : NIS yang baru saja disimpan (pencarian mulai sesudahnya); None = dari awal
    Bila semua sudah di-scan: mengembalikan `after` (bila ada di daftar) atau NIS pertama.
    """
    if not order:
        return None
    start = order.index(after) + 1 if after in order else 0
    for step in range(len(order)):
        nis = order[(start + step) % len(order)]
        if nis not in scanned:
            return nis
    return after if after in order else order[0]


# ---- kunci jawaban & bobot -------------------------------------------------
def parse_key_string(text, n):
    """'ABCDE...' (boleh ada spasi / koma / nomor) -> list huruf sepanjang n."""
    letters = [c for c in str(text or "").upper() if c.isalpha()]
    if not letters:
        raise ValueError("Isi dulu deretan huruf kunci jawaban.")
    invalid = sorted({c for c in letters if c not in CHOICES})
    if invalid:
        raise ValueError(f"Huruf tidak valid: {', '.join(invalid)}. Gunakan A-E saja.")
    if len(letters) != n:
        raise ValueError(f"Jumlah huruf ({len(letters)}) harus sama dengan jumlah soal ({n}).")
    return letters


def build_weights(mode, n, manual_values=None):
    """Bobot per soal sesuai mode. Melempar ValueError bila bobot manual tidak valid."""
    if mode == MODE_100:
        return [100.0 / n] * n
    if mode == MODE_1:
        return [1.0] * n
    if mode == MODE_MANUAL:
        values = []
        for i, v in enumerate(list(manual_values or []), start=1):
            if v is None or (isinstance(v, float) and v != v):
                raise ValueError(f"Bobot soal {i} masih kosong.")
            try:
                w = float(v)
            except (TypeError, ValueError):
                raise ValueError(f"Bobot soal {i} harus berupa angka.") from None
            if w < 0:
                raise ValueError(f"Bobot soal {i} tidak boleh negatif.")
            values.append(w)
        if len(values) != n:
            raise ValueError(f"Jumlah bobot ({len(values)}) harus sama dengan jumlah soal ({n}).")
        if sum(values) <= 0:
            raise ValueError("Total bobot harus lebih dari 0.")
        return values
    raise ValueError(f"Mode bobot tidak dikenal: {mode!r}")


# ---- PDF LJK -----------------------------------------------------------------
def make_pdf_bytes(quiz_name, total_questions, school, quiz_id, tanggal_iso):
    """Buat PDF LJK; Quiz_ID pada QR code diisi id kuis dari database."""
    with tempfile.TemporaryDirectory() as tmp:
        path = os.path.join(tmp, "ljk.pdf")
        gen.generate_ljk_pdf(
            quiz_name, total_questions, path,
            nama_sekolah=school, quiz_id=str(quiz_id), tanggal=tanggal_iso,
        )
        with open(path, "rb") as fh:
            return fh.read()


# ---- impor siswa ---------------------------------------------------------------
def import_students_bytes(data, filename):
    """Simpan unggahan ke file sementara lalu impor lewat database_manager."""
    ext = os.path.splitext(filename or "")[1].lower() or ".xlsx"
    with tempfile.TemporaryDirectory() as tmp:
        path = os.path.join(tmp, "siswa" + ext)
        with open(path, "wb") as fh:
            fh.write(data)
        return dm.import_students_from_excel(path)


def student_template_bytes():
    """Template Excel kosong: NIS, Nama_Lengkap, Kelas (kolom NIS berformat teks)."""
    from openpyxl import Workbook
    from openpyxl.styles import Font

    wb = Workbook()
    ws = wb.active
    ws.title = "Siswa"
    ws.append(["NIS", "Nama_Lengkap", "Kelas"])
    ws.append(["0001", "Contoh Siswa Satu", "X IPA 1"])
    ws.append(["0002", "Contoh Siswa Dua", "X IPA 1"])
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for row in ws.iter_rows(min_row=2, max_col=1):
        row[0].number_format = "@"
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 32
    ws.column_dimensions["C"].width = 14
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ---- scan & simpan ---------------------------------------------------------------
def details_df(grade):
    rows = []
    for d in grade["details"]:
        ans = d["jawaban"]
        if ans is None:
            status = "⚪ Kosong"
        elif len(ans) > 1:
            status = "⚠️ Ganda"
        elif d["benar"]:
            status = "✅ Benar"
        else:
            status = "❌ Salah"
        rows.append({
            "No": d["nomor"],
            "Jawaban": ans or "-",
            "Kunci": d["kunci"],
            "Status": status,
            "Bobot": round(d["bobot"], 2),
        })
    return pd.DataFrame(rows)


def run_scan(image_bytes, quiz, fill_threshold=DEFAULT_FILL_THRESHOLD, show_key=False):
    """Periksa satu foto LJK terhadap kuis. Melempar omr.OMRError / ValueError bila gagal."""
    img = cv2.imdecode(np.frombuffer(image_bytes, np.uint8), cv2.IMREAD_COLOR)
    if img is None:
        raise omr.OMRError("File bukan gambar yang bisa dibaca. Gunakan foto JPG / PNG.")

    h, w = img.shape[:2]
    notes = []
    if min(h, w) < LOW_RES_SIDE:
        notes.append(
            f"Foto beresolusi rendah ({w}x{h}). Pembacaan bisa gagal atau kurang akurat. "
            "Gunakan 📁 Unggah foto (kamera bawaan HP, resolusi penuh) atau dekatkan kertas "
            "sampai memenuhi layar."
        )
    big = max(h, w)
    if big > MAX_IMAGE_SIDE:
        f = MAX_IMAGE_SIDE / big
        img = cv2.resize(img, (int(w * f), int(h * f)), interpolation=cv2.INTER_AREA)

    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        res = omr.scan_ljk(
            img, quiz["key"], total_questions=quiz["total_questions"],
            weights=quiz["weights"], show_correct_key=show_key,
            fill_threshold=fill_threshold,
        )
    notes += [str(m.message) for m in caught if issubclass(m.category, RuntimeWarning)]

    meta = res["metadata"]
    mismatch = None
    if meta:
        if str(meta.get("Quiz_ID")) != str(quiz["id"]):
            mismatch = (
                f"QR pada LJK ini menunjuk kuis ID '{meta.get('Quiz_ID')}', sedangkan kuis yang "
                f"dipilih ber-ID '{quiz['id']}'. Kemungkinan LJK berasal dari kuis lain."
            )
    else:
        notes.append("QR code tidak terbaca, jadi kecocokan LJK dengan kuis tidak bisa diperiksa.")

    return {
        "answers": res["student_answers"],
        "grade": res["grade"],
        "feedback_bgr": res["feedback"],
        "metadata": meta,
        "alignment_score": res["alignment_score"],
        "notes": notes,
        "mismatch": mismatch,
        "image_size": (w, h),
    }


def resize_to_width(img_bgr, width):
    h, w = img_bgr.shape[:2]
    if w <= width:
        return img_bgr
    return cv2.resize(img_bgr, (width, int(h * width / w)), interpolation=cv2.INTER_AREA)


def jpeg_bytes(img_bgr, quality=85):
    ok, buf = cv2.imencode(".jpg", img_bgr, [cv2.IMWRITE_JPEG_QUALITY, quality])
    return buf.tobytes() if ok else b""


def save_scan(scan, quiz, nis, photo_bytes, photo_name="foto.jpg", scans_dir=None):
    """Simpan foto asli + gambar koreksi (JPG kecil) ke disk lalu hasilnya ke database.

    Mengembalikan id scan.
    """
    if scans_dir is None:
        scans_dir = os.path.join(os.path.dirname(os.path.abspath(dm.DB_PATH)), "scans")
    folder = os.path.join(scans_dir, f"kuis_{quiz['id']}")
    os.makedirs(folder, exist_ok=True)

    ext = os.path.splitext(photo_name or "")[1].lower()
    if ext not in (".jpg", ".jpeg", ".png"):
        ext = ".jpg"
    base = f"{_safe_name(nis)}_{datetime.now():%Y%m%d_%H%M%S_%f}"
    orig_path = os.path.join(folder, base + ext)
    fb_path = os.path.join(folder, base + "_koreksi.jpg")

    written = []
    try:
        with open(orig_path, "wb") as fh:
            fh.write(photo_bytes)
        written.append(orig_path)
        data = jpeg_bytes(resize_to_width(scan["feedback_bgr"], 1000))
        if data:
            with open(fb_path, "wb") as fh:
                fh.write(data)
            written.append(fb_path)
        return dm.save_scan_result(
            quiz["id"], nis, scan["grade"]["total_score"], scan["answers"], orig_path
        )
    except Exception:
        for path in written:
            try:
                os.remove(path)
            except OSError:
                pass
        raise


def undo_scan(scan_id):
    """Batalkan satu hasil scan: hapus baris database dan file fotonya. True bila ada."""
    rows = _rows("SELECT image_path FROM scan_results WHERE id = ?", (int(scan_id),))
    if not rows:
        return False
    ok = dm.delete_scan_result(scan_id)
    path = rows[0]["image_path"]
    if ok and path:
        base, _ = os.path.splitext(path)
        for p in (path, base + "_koreksi.jpg", base + "_koreksi.png"):
            try:
                os.remove(p)
            except OSError:
                pass
    return ok


# ---- laporan -----------------------------------------------------------------------
def _latest_rows(quiz_id):
    return _rows(
        """
        SELECT s.student_id AS nis, s.full_name AS nama, c.class_name AS kelas,
               r.total_score AS total_score, r.student_answers_json AS answers_json,
               r.scanned_at AS scanned_at
        FROM scan_results r
        JOIN students s ON s.id = r.student_id
        JOIN classes  c ON c.id = s.class_id
        WHERE r.quiz_id = ?
          AND r.id = (SELECT MAX(r2.id) FROM scan_results r2
                      WHERE r2.quiz_id = r.quiz_id AND r2.student_id = r.student_id)
        ORDER BY c.class_name COLLATE NOCASE, s.full_name COLLATE NOCASE
        """,
        (quiz_id,),
    )


RESULT_COLUMNS = ["NIS", "Nama", "Kelas", "Total Skor", "Benar", "Salah", "Persentase", "Waktu Scan"]


def results_df(quiz_id):
    """Nilai terbaru tiap siswa (kolom sama dengan ekspor Excel + waktu scan)."""
    quiz = get_quiz(quiz_id)
    n, keys, max_score = quiz["total_questions"], quiz["key"], quiz["max_score"]
    records = []
    for r in _latest_rows(quiz_id):
        answers = json.loads(r["answers_json"])
        correct = sum(1 for i, k in enumerate(keys, start=1) if answers.get(str(i)) == k)
        records.append([
            r["nis"], r["nama"], r["kelas"], r["total_score"], correct, n - correct,
            round(r["total_score"] / max_score * 100, 2) if max_score > 0 else 0.0,
            r["scanned_at"],
        ])
    return pd.DataFrame(records, columns=RESULT_COLUMNS)


def score_stats(df, kkm):
    p = df["Persentase"]
    return {
        "peserta": int(len(df)),
        "rata": round(float(p.mean()), 2),
        "median": round(float(p.median()), 2),
        "maks": round(float(p.max()), 2),
        "min": round(float(p.min()), 2),
        "std": round(float(p.std(ddof=0)), 2),
        "lulus": int((p >= kkm).sum()),
    }


def score_distribution(df):
    labels = [f"{a:02d}-{a + 9:02d}" for a in range(0, 90, 10)] + ["90-100"]
    counts = [0] * 10
    for p in df["Persentase"]:
        counts[min(int(p // 10), 9)] += 1
    return pd.DataFrame({"Jumlah siswa": counts}, index=labels)


def class_summary(df):
    return (
        df.groupby("Kelas")
        .agg(Peserta=("NIS", "count"), Rata=("Persentase", "mean"),
             Tertinggi=("Persentase", "max"), Terendah=("Persentase", "min"))
        .round(2)
        .reset_index()
        .rename(columns={"Rata": "Rata-rata"})
    )


def item_analysis(quiz_id):
    """Tingkat keberhasilan per soal dari scan terbaru tiap siswa."""
    quiz = get_quiz(quiz_id)
    rows = _latest_rows(quiz_id)
    n = quiz["total_questions"]
    correct, blank = [0] * n, [0] * n
    for r in rows:
        answers = json.loads(r["answers_json"])
        for i, k in enumerate(quiz["key"], start=1):
            a = answers.get(str(i))
            if a == k:
                correct[i - 1] += 1
            elif a is None:
                blank[i - 1] += 1
    total = max(len(rows), 1)
    return pd.DataFrame({
        "No": list(range(1, n + 1)),
        "Kunci": quiz["key"],
        "Benar": correct,
        "Kosong": blank,
        "% Benar": [round(c / total * 100, 1) for c in correct],
    })


def export_bytes(quiz_id, fmt="xlsx"):
    """Ekspor nilai lewat database_manager. Mengembalikan (bytes, jumlah_baris)."""
    with tempfile.TemporaryDirectory() as tmp:
        path = os.path.join(tmp, f"nilai.{fmt}")
        rows = dm.export_grades_to_excel(quiz_id, path)
        with open(path, "rb") as fh:
            return fh.read(), rows


def check_password(entered, expected):
    """Bandingkan kata sandi dengan waktu konstan."""
    return hmac.compare_digest(
        hashlib.sha256(str(entered).encode("utf-8")).digest(),
        hashlib.sha256(str(expected).encode("utf-8")).digest(),
    )


def configured_password():
    """Kata sandi aplikasi dari Streamlit secrets / environment variable (None = tanpa sandi)."""
    value = None
    try:
        value = st.secrets.get("APP_PASSWORD")
    except Exception:  # tidak ada file secrets
        value = None
    value = value or os.environ.get("APP_PASSWORD")
    return str(value) if value else None


def app_icon_png_b64(size=192):
    """Ikon aplikasi sederhana (PNG, base64) untuk 'Tambahkan ke Layar Utama'."""
    img = np.zeros((size, size, 3), np.uint8)
    img[:] = (235, 99, 37)  # BGR dari #2563eb
    c = size // 2
    cv2.circle(img, (c, c), int(size * 0.30), (255, 255, 255), max(4, size // 24), cv2.LINE_AA)
    cv2.circle(img, (c, c), int(size * 0.13), (255, 255, 255), -1, cv2.LINE_AA)
    ok, buf = cv2.imencode(".png", img)
    return base64.b64encode(buf.tobytes()).decode("ascii") if ok else ""


# ==========================================================================
# TAMPILAN: CSS, komponen kecil
# ==========================================================================
CSS = """
<style>
:root{
  --brand:#2563eb; --brand-dark:#1d4ed8; --ok:#16a34a; --bad:#dc2626; --warn:#d97706;
  --bg:#f3f5f9; --card:#ffffff; --ink:#0f172a; --muted:#64748b; --line:#e2e8f0;
}
html, body, [data-testid="stApp"], [data-testid="stAppViewContainer"]{ background:var(--bg); }
header[data-testid="stHeader"], [data-testid="stToolbar"], [data-testid="stDecoration"],
[data-testid="stStatusWidget"], .stAppDeployButton, #MainMenu, footer{ display:none !important; }

/* Lebar seperti layar HP, ruang bawah untuk bottom navbar */
.block-container{ max-width:560px !important; padding:.9rem .9rem 7.5rem !important; }
h1{ font-size:1.55rem !important; margin:.2rem 0 .6rem !important; }
h2, h3{ font-size:1.15rem !important; }

/* ===== Kartu ===== */
div[data-testid="stVerticalBlockBorderWrapper"]{
  background:var(--card); border:1px solid var(--line) !important; border-radius:18px !important;
  box-shadow:0 2px 10px rgba(15,23,42,.05); padding:.35rem .25rem;
}
.hero{ background:linear-gradient(135deg,var(--brand),#38bdf8); color:#fff; border-radius:20px;
  padding:16px 18px; margin-bottom:12px; }
.hero .t{ font-size:1.25rem; font-weight:800; } .hero .s{ opacity:.9; font-size:.9rem; margin-top:2px; }
.mgrid{ display:grid; grid-template-columns:repeat(auto-fit,minmax(96px,1fr)); gap:8px; margin:6px 0 12px; }
.mcell{ background:var(--card); border:1px solid var(--line); border-radius:14px; padding:10px 6px; text-align:center; }
.mcell .v{ font-size:1.35rem; font-weight:800; color:var(--ink); line-height:1.1; }
.mcell .l{ font-size:.72rem; color:var(--muted); margin-top:3px; }
.mcell.ok .v{ color:var(--ok);} .mcell.bad .v{ color:var(--bad);} .mcell.warn .v{ color:var(--warn);}
.score{ text-align:center; padding:6px 0 2px; }
.score .big{ font-size:2.6rem; font-weight:900; color:var(--brand); line-height:1; }
.score .sub{ color:var(--muted); font-size:.95rem; margin-top:4px; }
.who{ font-weight:700; font-size:1.05rem; } .meta{ color:var(--muted); font-size:.85rem; }
.pill{ display:inline-block; padding:2px 10px; border-radius:999px; font-size:.78rem; font-weight:700;
  background:#e0e7ff; color:#3730a3; }
.pill.ok{ background:#dcfce7; color:#166534; } .pill.bad{ background:#fee2e2; color:#991b1b; }

/* ===== Tombol besar (touch-friendly) ===== */
.stButton > button, .stDownloadButton > button, [data-testid="stFormSubmitButton"] > button{
  width:100%; min-height:3.4rem; border-radius:16px; font-weight:700; font-size:1.05rem;
  border:1px solid var(--line); transition:transform .05s;
}
.stButton > button:active, .stDownloadButton > button:active{ transform:scale(.98); }
.stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"]{
  background:var(--brand); border-color:var(--brand); color:#fff; box-shadow:0 4px 12px rgba(37,99,235,.3);
}
.stButton > button[kind="primary"]:hover{ background:var(--brand-dark); border-color:var(--brand-dark); color:#fff; }

/* ===== Input besar, font 16px (tidak memicu zoom iOS) ===== */
input, textarea, [data-baseweb="select"] > div{ font-size:16px !important; }
[data-baseweb="input"], [data-baseweb="select"] > div, [data-baseweb="textarea"]{ border-radius:12px !important; }
[data-testid="stNumberInput"] input{ min-height:2.9rem; }
label p{ font-weight:600; }

/* ===== Kamera: selebar layar, tombol jepret besar di bawah ===== */
[data-testid="stCameraInput"]{ width:calc(100% + 1.8rem); margin-left:-.9rem; margin-right:-.9rem; }
[data-testid="stCameraInput"] video, [data-testid="stCameraInput"] img{
  width:100% !important; height:auto !important; border-radius:0; display:block; }
[data-testid="stCameraInput"] > div{ width:100%; }
[data-testid="stCameraInputButton"]{
  width:calc(100% - 1.8rem) !important; margin:10px .9rem 4px !important; min-height:4.4rem !important;
  border-radius:22px !important; font-size:1.25rem !important; font-weight:800 !important;
  background:var(--brand) !important; color:#fff !important; border:0 !important;
  box-shadow:0 6px 18px rgba(37,99,235,.38) !important;
}
[data-testid="stImage"] img{ width:100% !important; border-radius:12px; }

/* ===== Bottom Navigation Bar ===== */
.st-key-bottom_nav{
  position:fixed; left:50%; bottom:0; transform:translateX(-50%); width:min(100%,560px); z-index:1000;
  background:rgba(255,255,255,.97); backdrop-filter:blur(10px); border-top:1px solid var(--line);
  padding:6px 8px calc(6px + env(safe-area-inset-bottom)); box-shadow:0 -6px 20px rgba(15,23,42,.08);
}
.st-key-bottom_nav [role="radiogroup"]{ display:flex; gap:4px; width:100%; }
.st-key-bottom_nav label{
  flex:1 1 0; margin:0 !important; padding:10px 2px !important; border-radius:14px; cursor:pointer;
  justify-content:center; text-align:center; min-height:3.2rem; align-items:center;
}
.st-key-bottom_nav label > div:first-child{ display:none !important; }
.st-key-bottom_nav label div[data-testid="stMarkdownContainer"] p{
  font-size:.86rem; font-weight:700; white-space:nowrap; margin:0; color:var(--muted); }
.st-key-bottom_nav label:has(input:checked){ background:var(--brand); }
.st-key-bottom_nav label:has(input:checked) p{ color:#fff !important; }
.st-key-bottom_nav [data-testid="stWidgetLabel"]{ display:none; }
</style>
"""


def inject_css():
    st.markdown(CSS, unsafe_allow_html=True)


def inject_pwa():
    """Meta tag & manifest agar 'Tambahkan ke Layar Utama' tampil seperti aplikasi.

    Streamlit Cloud tidak mengizinkan service worker / manifest di root, jadi ini upaya
    terbaik: tanpa mode offline, tetapi bisa dibuka layar penuh dari ikon di layar utama.
    """
    try:
        import streamlit.components.v1 as components

        icon = "data:image/png;base64," + app_icon_png_b64()
        script = """
<script>
try {
  const d = window.parent.document, h = d.head;
  const meta = (name, content) => {
    let m = h.querySelector('meta[name="' + name + '"]');
    if (!m) { m = d.createElement('meta'); m.setAttribute('name', name); h.appendChild(m); }
    m.setAttribute('content', content);
  };
  meta('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
  meta('theme-color', '#2563eb');
  meta('mobile-web-app-capable', 'yes');
  meta('apple-mobile-web-app-capable', 'yes');
  meta('apple-mobile-web-app-title', 'LJK');
  meta('apple-mobile-web-app-status-bar-style', 'default');
  const icon = __ICON__;
  const link = (rel, href, extra) => {
    let l = h.querySelector('link[rel="' + rel + '"]');
    if (!l) { l = d.createElement('link'); l.setAttribute('rel', rel); h.appendChild(l); }
    l.setAttribute('href', href);
    if (extra) for (const k in extra) l.setAttribute(k, extra[k]);
  };
  link('apple-touch-icon', icon);
  const manifest = {name: 'Pemeriksa LJK', short_name: 'LJK', display: 'standalone',
    background_color: '#f3f5f9', theme_color: '#2563eb',
    start_url: window.parent.location.pathname || '/',
    icons: [{src: icon, sizes: '192x192', type: 'image/png'}]};
  link('manifest', 'data:application/manifest+json,' + encodeURIComponent(JSON.stringify(manifest)));
} catch (e) {}
</script>
""".replace("__ICON__", json.dumps(icon))
        components.html(script, height=0)
    except Exception:
        pass  # PWA hanyalah pemanis; aplikasi tetap jalan tanpa ini


def esc(text):
    return html.escape(str(text))


def metric_grid(items):
    """Deretan angka ringkas. items: [(label, nilai, kelas_css_opsional), ...]"""
    cells = "".join(
        f'<div class="mcell {c[2] if len(c) > 2 else ""}"><div class="v">{esc(c[1])}</div>'
        f'<div class="l">{esc(c[0])}</div></div>'
        for c in items
    )
    st.markdown(f'<div class="mgrid">{cells}</div>', unsafe_allow_html=True)


def flash(level, text):
    """Pesan yang tetap tampil setelah st.rerun()."""
    st.session_state["flash"] = (level, text)


def show_flash():
    item = st.session_state.pop("flash", None)
    if not item:
        return
    level, text = item
    {"success": st.success, "error": st.error, "warning": st.warning}.get(level, st.info)(text)


def goto(page, quiz_id=None):
    """Callback tombol: pindah menu (boleh menetapkan kuis aktif). Aman mengubah state widget."""
    st.session_state["nav"] = page
    if quiz_id is not None:
        st.session_state["active_quiz"] = quiz_id
        st.session_state["scan_quiz"] = quiz_id


@st.cache_data(show_spinner=False, max_entries=20)
def cached_pdf(quiz_name, total_questions, school, quiz_id, tanggal_iso):
    return make_pdf_bytes(quiz_name, total_questions, school, quiz_id, tanggal_iso)


def quiz_label(q):
    return f"{q['name']} · {q['total_questions']} soal"


# ==========================================================================
# HALAMAN 🏠 KUIS
# ==========================================================================
def page_kuis():
    # Form dikosongkan SETELAH simpan (di awal run berikutnya, sebelum widget dibuat).
    if st.session_state.pop("reset_form", False):
        st.session_state["new_name"] = ""
        st.session_state["quick_key"] = ""
        st.session_state["w_mode"] = MODE_100

    counts = get_counts()
    st.markdown(
        '<div class="hero"><div class="t">Pemeriksa LJK</div>'
        '<div class="s">Buat kuis → cetak LJK → scan dengan HP → rekap nilai</div></div>',
        unsafe_allow_html=True,
    )
    metric_grid([("Kuis", counts["kuis"]), ("Siswa", counts["siswa"]),
                 ("LJK ter-scan", counts["scan"])])
    show_flash()

    # ---- Kuis baru ----
    with st.container(border=True):
        st.subheader("➕ Kuis baru")
        name = st.text_input("Nama kuis", key="new_name", placeholder="mis. Kuis Matematika Bab 3")
        n = int(st.number_input("Jumlah soal", min_value=MIN_SOAL, max_value=MAX_SOAL,
                                value=30, step=1, key="new_n",
                                help="Boleh berapa pun dari 1 sampai 100 (mis. 35 atau 50)."))
        quick = st.text_area(
            "⚡ Isi Cepat Kunci Jawaban",
            key="quick_key", height=110,
            placeholder="Tempel sekali, mis.  ABCDEABCDE…  atau  1.A 2.B 3.C …",
            help="Huruf A-E saja; spasi, koma, baris baru, dan nomor diabaikan.",
        )
        letters = [c for c in quick.upper() if c.isalpha()]
        invalid = sorted({c for c in letters if c not in CHOICES})
        if not letters:
            st.caption(f"Belum ada kunci. Butuh {n} huruf (A-E).")
        elif invalid:
            st.markdown(f":red[Huruf tidak valid: {esc(', '.join(invalid))}. Gunakan A-E saja.]")
        elif len(letters) == n:
            st.markdown(f":green[✔ {n} dari {n} huruf. Kunci lengkap.]")
        else:
            st.markdown(f":orange[⚠ {len(letters)} dari {n} huruf.]")
        if letters and not invalid:
            preview = "  ".join(f"{i}{c}" for i, c in enumerate(letters[:MAX_SOAL], start=1))
            st.caption(preview)

        with st.expander("⚖️ Bobot nilai"):
            mode = st.radio("Mode bobot", WEIGHT_MODES, key="w_mode")
            manual_values = None
            if mode == MODE_MANUAL:
                w_df = pd.DataFrame({"No": list(range(1, n + 1)), "Bobot": [1.0] * n})
                w_edit = st.data_editor(
                    w_df, key=f"w_editor_{n}", hide_index=True, disabled=["No"],
                    num_rows="fixed", height=320,
                    column_config={"Bobot": st.column_config.NumberColumn(
                        "Bobot", min_value=0.0, step=0.5, format="%.2f")},
                )
                manual_values = [None if pd.isna(v) else float(v) for v in w_edit["Bobot"]]
            try:
                weights = build_weights(mode, n, manual_values)
            except ValueError as exc:
                weights = None
                st.error(str(exc))
            else:
                st.caption(f"Skor maksimum kuis: **{sum(weights):g}**")

        if st.button("💾 Simpan kuis", type="primary", key="save_quiz"):
            problems = []
            if not (name or "").strip():
                problems.append("Nama kuis wajib diisi.")
            keys = None
            try:
                keys = parse_key_string(quick, n)
            except ValueError as exc:
                problems.append(str(exc))
            if weights is None:
                problems.append("Bobot nilai belum valid.")
            if problems:
                for msg in problems:
                    st.error(msg)
            else:
                qid = dm.create_quiz(name.strip(), n, keys, weights)
                st.session_state["active_quiz"] = qid
                st.session_state["reset_form"] = True
                flash("success", f"Kuis '{name.strip()}' tersimpan. Unduh LJK-nya di daftar bawah, "
                                 "lalu tekan 📸 Scan untuk mulai memeriksa.")
                st.rerun()

    # ---- Daftar kuis ----
    quizzes = list_quizzes()
    st.subheader("Daftar kuis")
    if not quizzes:
        st.info("Belum ada kuis. Buat kuis pertama di atas.")
    with st.expander("🖨️ Opsi cetak LJK", expanded=False):
        st.text_input("Nama sekolah (tercetak di LJK)", key="school_name", value="Nama Sekolah")
    school = st.session_state.get("school_name") or "Nama Sekolah"

    for q in quizzes:
        with st.container(border=True):
            active = st.session_state.get("active_quiz") == q["id"]
            st.markdown(
                f'<div class="who">{esc(q["name"])} {"<span class=pill>aktif</span>" if active else ""}</div>'
                f'<div class="meta">{q["total_questions"]} soal · {q["peserta"]} siswa ter-scan · ID {q["id"]}</div>',
                unsafe_allow_html=True,
            )
            st.button("📸 Mulai scan kuis ini", key=f"scan_{q['id']}", type="primary",
                      on_click=goto, args=(NAV_SCAN, q["id"]))
            if st.session_state.get("pdf_for") == q["id"]:
                pdf = cached_pdf(q["name"], q["total_questions"], school, q["id"],
                                 date.today().isoformat())
                st.download_button("⬇️ Unduh PDF LJK", pdf,
                                   file_name=f"LJK_{_safe_name(q['name'])}_{q['total_questions']}soal.pdf",
                                   mime="application/pdf", key=f"dl_pdf_{q['id']}")
            else:
                if st.button("🖨️ Siapkan PDF LJK", key=f"mk_pdf_{q['id']}"):
                    st.session_state["pdf_for"] = q["id"]
                    st.rerun()
            if st.session_state.get("del_quiz") == q["id"]:
                st.warning("Hapus kuis ini beserta SEMUA nilainya? Tidak bisa dibatalkan.")
                if st.button("Ya, hapus", key=f"del_yes_{q['id']}", type="primary"):
                    dm.delete_quiz(q["id"])
                    for k in ("del_quiz", "pdf_for"):
                        st.session_state.pop(k, None)
                    if st.session_state.get("active_quiz") == q["id"]:
                        st.session_state.pop("active_quiz", None)
                    flash("success", f"Kuis '{q['name']}' dihapus.")
                    st.rerun()
                if st.button("Batal", key=f"del_no_{q['id']}"):
                    st.session_state.pop("del_quiz", None)
                    st.rerun()
            else:
                if st.button("🗑️ Hapus kuis", key=f"del_{q['id']}"):
                    st.session_state["del_quiz"] = q["id"]
                    st.rerun()

    # ---- Cadangan ----
    with st.expander("💾 Cadangan & pulihkan data"):
        st.caption(
            "Di Streamlit Cloud, data bisa hilang saat aplikasi restart atau tidur. "
            "Unduh cadangan secara rutin (mis. setiap selesai scan) dan pulihkan bila perlu."
        )
        st.download_button("⬇️ Unduh cadangan data", dm.backup_bytes(),
                           file_name=f"cadangan_ljk_{datetime.now():%Y%m%d_%H%M}.db",
                           mime="application/octet-stream", key="dl_backup")
        up = st.file_uploader("Pulihkan dari cadangan (.db)", type=["db", "sqlite"], key="restore_up")
        agree = st.checkbox("Saya mengerti: data sekarang akan DIGANTI oleh cadangan ini",
                            key="restore_ok")
        if st.button("♻️ Pulihkan data", key="restore_btn", disabled=not (up and agree)):
            try:
                dm.restore_from_bytes(up.getvalue())
            except ValueError as exc:
                st.error(str(exc))
            else:
                st.session_state.pop("active_quiz", None)
                flash("success", "Data berhasil dipulihkan dari cadangan.")
                st.rerun()


# ==========================================================================
# HALAMAN 📸 SCAN
# ==========================================================================
def _cb_bump_camera():
    st.session_state["cam_ver"] = st.session_state.get("cam_ver", 0) + 1


def _cb_undo(scan_id, nis):
    if undo_scan(scan_id):
        flash("success", "Penyimpanan terakhir dibatalkan.")
        st.session_state["scan_stu"] = nis
    st.session_state.pop("last_saved", None)


def page_scan():
    show_flash()
    quizzes = list_quizzes()
    if not quizzes:
        st.info("Belum ada kuis. Buat kuis dulu di menu 🏠 Kuis.")
        st.button("🏠 Ke menu Kuis", key="to_kuis", type="primary", on_click=goto, args=(NAV_KUIS,))
        return
    ids = [q["id"] for q in quizzes]
    labels = {q["id"]: quiz_label(q) for q in quizzes}
    if st.session_state.get("scan_quiz") not in ids:
        act = st.session_state.get("active_quiz")
        st.session_state["scan_quiz"] = act if act in ids else ids[0]
    qid = st.selectbox("Kuis", ids, key="scan_quiz", format_func=lambda i: labels[i])
    st.session_state["active_quiz"] = qid
    quiz = get_quiz(qid)

    students = list_students()
    if students.empty:
        st.warning("Belum ada data siswa. Impor daftar siswa dulu agar nilai bisa disimpan.")
        st.button("👥 Ke menu Siswa", key="to_siswa", type="primary", on_click=goto, args=(NAV_SISWA,))
        return

    classes = sorted(students["Kelas"].unique(), key=str.lower)
    if len(classes) > 1:
        if st.session_state.get("scan_kelas") not in [ALL_CLASSES] + classes:
            st.session_state["scan_kelas"] = ALL_CLASSES
        kelas = st.selectbox("Kelas", [ALL_CLASSES] + classes, key="scan_kelas")
        subset = students if kelas == ALL_CLASSES else students[students["Kelas"] == kelas]
    else:
        subset = students
    order = list(subset["NIS"])
    info = {r.NIS: (r.Nama, r.Kelas) for r in students.itertuples(index=False)}
    scanned = scanned_nis(qid)

    # Lompat ke siswa berikutnya yang belum di-scan (sebelum widget siswa dibuat)
    advance = st.session_state.pop("advance_after", None)
    if advance is not None or st.session_state.get("scan_stu") not in order:
        st.session_state["scan_stu"] = next_unscanned(order, scanned, advance)

    def stu_label(nis):
        nama, kls = info[nis]
        return f"{'✅ ' if nis in scanned else ''}{nama} · {kls}"

    nis = st.selectbox("Siswa", order, key="scan_stu", format_func=stu_label)
    nama, kls = info[nis]
    done = len(scanned & set(order))
    st.caption(f"Sudah di-scan: {done} dari {len(order)} siswa" +
               (" · semua selesai 🎉" if done == len(order) else ""))
    if nis in scanned:
        st.markdown(":orange[Siswa ini sudah punya nilai. Menyimpan lagi akan menggantinya "
                    "sebagai nilai terbaru.]")

    with st.expander("⚙️ Opsi scan"):
        auto_save = st.checkbox("Simpan otomatis begitu LJK terbaca", key="auto_save",
                                help="Tanpa menekan tombol simpan; langsung lanjut ke siswa berikutnya. "
                                     "Hasil terakhir tetap tampil dan bisa dibatalkan.")
        threshold = st.slider("Sensitivitas arsiran (batas isi bulatan)", 0.20, 0.70,
                              DEFAULT_FILL_THRESHOLD, 0.05, key="fill_thr",
                              help="Naikkan bila bulatan kosong terbaca terisi; turunkan bila "
                                   "arsiran tipis tidak terbaca.")
        show_key = st.checkbox("Tampilkan kunci benar pada soal yang salah (lingkaran hijau tipis)",
                               key="show_key")

    src = st.radio("Sumber foto", [SRC_CAMERA, SRC_UPLOAD], key="scan_src", horizontal=True,
                   label_visibility="collapsed")
    ver = st.session_state.setdefault("cam_ver", 0)
    if src == SRC_CAMERA:
        photo = st.camera_input("Jepret LJK", key=f"cam_{ver}", label_visibility="collapsed")
        st.caption("Pegang HP tegak, pastikan 4 kotak hitam sudut terlihat utuh. Resolusi kamera "
                   "browser cukup rendah; untuk hasil paling akurat pakai 📁 Unggah foto.")
    else:
        photo = st.file_uploader("Ambil / pilih foto LJK", type=["jpg", "jpeg", "png"],
                                 key=f"up_{ver}", label_visibility="collapsed")
        st.caption("Di HP, tombol ini bisa membuka kamera bawaan dengan resolusi penuh.")

    if photo is not None:
        data = photo.getvalue()
        _render_scan_result(quiz, nis, nama, kls, data, getattr(photo, "name", "foto.jpg"),
                            threshold, show_key, auto_save)

    last = st.session_state.get("last_saved")
    if last:
        _render_last_saved(last)


def _render_scan_result(quiz, nis, nama, kls, data, photo_name, threshold, show_key, auto_save):
    sig = (hashlib.md5(data).hexdigest(), quiz["id"], round(threshold, 3), bool(show_key))
    cache = st.session_state.get("scan_cache")
    if cache and cache["sig"] == sig:
        scan, error = cache["scan"], cache["error"]
    else:
        scan, error = None, None
        with st.spinner("Memeriksa LJK…"):
            try:
                scan = run_scan(data, quiz, threshold, show_key)
            except (omr.OMRError, ValueError) as exc:
                error = str(exc)
        st.session_state["scan_cache"] = {"sig": sig, "scan": scan, "error": error}

    if error:
        with st.container(border=True):
            st.error(error)
            st.caption("Tips: cahaya merata tanpa bayangan, kertas rata, seluruh kertas dan 4 kotak "
                       "hitam sudut terlihat, jumlah soal kuis sama dengan LJK.")
            st.button("🔁 Foto ulang", key="retry_err", type="primary", on_click=_cb_bump_camera)
        return

    grade = scan["grade"]
    with st.container(border=True):
        st.markdown(f'<div class="who">{esc(nama)}</div><div class="meta">{esc(kls)} · NIS {esc(nis)}</div>',
                    unsafe_allow_html=True)
        st.markdown(
            f'<div class="score"><div class="big">{grade["total_score"]:g}</div>'
            f'<div class="sub">dari {grade["max_score"]:g} · {grade["persentase"]:g}%</div></div>',
            unsafe_allow_html=True,
        )
        salah_murni = max(grade["salah"] - grade["kosong"] - grade["ganda"], 0)
        metric_grid([("Benar", grade["benar"], "ok"), ("Salah", salah_murni, "bad"),
                     ("Kosong", grade["kosong"]), ("Ganda", grade["ganda"], "warn")])
        if scan["mismatch"]:
            st.warning(scan["mismatch"])
        for note in scan["notes"]:
            st.markdown(f":orange[⚠ {esc(note)}]")
        st.image(cv2.cvtColor(resize_to_width(scan["feedback_bgr"], 1000), cv2.COLOR_BGR2RGB),
                 caption="Hijau = benar · Merah = salah")
        with st.expander("Detail jawaban per soal"):
            st.dataframe(details_df(grade), hide_index=True)

        blocked = bool(scan["mismatch"])
        if st.button("✅ Simpan & lanjut ke siswa berikutnya", key="save_next", type="primary"):
            _do_save(scan, quiz, nis, nama, kls, data, photo_name)
        st.button("🔁 Foto ulang (jangan simpan)", key="retry_ok", on_click=_cb_bump_camera)

    if auto_save and not blocked and st.session_state.get("auto_done") != sig:
        st.session_state["auto_done"] = sig
        _do_save(scan, quiz, nis, nama, kls, data, photo_name)


def _do_save(scan, quiz, nis, nama, kls, data, photo_name):
    try:
        sid = save_scan(scan, quiz, nis, data, photo_name)
    except Exception as exc:  # noqa: BLE001 - tampilkan apa pun penyebabnya ke pengguna
        st.error(f"Gagal menyimpan: {exc}")
        return
    grade = scan["grade"]
    st.session_state["last_saved"] = {
        "id": sid, "nis": nis, "nama": nama, "kelas": kls,
        "skor": grade["total_score"], "maks": grade["max_score"], "persen": grade["persentase"],
        "thumb": jpeg_bytes(resize_to_width(scan["feedback_bgr"], 700), 80),
    }
    st.session_state["advance_after"] = nis
    st.session_state["cam_ver"] = st.session_state.get("cam_ver", 0) + 1
    st.session_state.pop("scan_cache", None)
    flash("success", f"✅ {nama}: {grade['total_score']:g}/{grade['max_score']:g} tersimpan.")
    st.rerun()


def _render_last_saved(last):
    with st.container(border=True):
        st.markdown(
            f'<div class="meta">Terakhir disimpan</div><div class="who">{esc(last["nama"])} · '
            f'{last["skor"]:g}/{last["maks"]:g} ({last["persen"]:g}%)</div>',
            unsafe_allow_html=True,
        )
        if last.get("thumb"):
            st.image(last["thumb"])
        st.button("↩️ Batalkan simpan ini", key="undo_last", on_click=_cb_undo,
                  args=(last["id"], last["nis"]))


# ==========================================================================
# HALAMAN 👥 SISWA
# ==========================================================================
def page_siswa():
    # Isian dikosongkan SETELAH simpan, di awal run berikutnya (sebelum widget dibuat).
    if st.session_state.pop("reset_add", False):
        st.session_state["add_nis"] = ""
        st.session_state["add_nama"] = ""
    show_flash()
    st.title("👥 Data Siswa")
    students = list_students()
    metric_grid([("Siswa", len(students)), ("Kelas", students["Kelas"].nunique() if len(students) else 0)])

    with st.container(border=True):
        st.subheader("📥 Impor dari Excel")
        st.caption("Kolom: NIS, Nama_Lengkap, Kelas (.xlsx atau .csv). NIS yang sudah ada akan diperbarui.")
        st.download_button("⬇️ Unduh template Excel", student_template_bytes(),
                           file_name="template_siswa.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           key="dl_tpl")
        up = st.file_uploader("Pilih file siswa", type=["xlsx", "csv"], key="stu_up")
        if st.button("📥 Impor sekarang", key="stu_import", type="primary", disabled=up is None):
            try:
                res = import_students_bytes(up.getvalue(), up.name)
            except Exception as exc:  # noqa: BLE001
                st.error(f"Gagal impor: {exc}")
            else:
                msg = f"Baru: {res['inserted']} · Diperbarui: {res['updated']} · Dilewati: {res['skipped']}"
                if res.get("errors"):
                    msg += "\n\n" + "\n".join(f"- {e}" for e in res["errors"][:10])
                flash("success" if not res.get("errors") else "warning", msg)
                st.rerun()

    with st.container(border=True):
        st.subheader("➕ Tambah satu siswa")
        c_nis = st.text_input("NIS", key="add_nis")
        c_nama = st.text_input("Nama lengkap", key="add_nama")
        c_kelas = st.text_input("Kelas", key="add_kelas", placeholder="mis. X IPA 1")
        if st.button("Simpan siswa", key="add_btn"):
            try:
                status = dm.add_student(c_nis, c_nama, c_kelas)
            except ValueError as exc:
                st.error(str(exc))
            else:
                st.session_state["reset_add"] = True
                flash("success", f"Siswa {c_nama.strip()} ({status}).")
                st.rerun()

    st.subheader("Daftar siswa")
    if students.empty:
        st.info("Belum ada siswa.")
        return
    classes = [ALL_CLASSES] + sorted(students["Kelas"].unique(), key=str.lower)
    kelas = st.selectbox("Tampilkan kelas", classes, key="stu_filter")
    view = students if kelas == ALL_CLASSES else students[students["Kelas"] == kelas]
    st.dataframe(view, hide_index=True)


# ==========================================================================
# HALAMAN 📊 REKAP
# ==========================================================================
def page_rekap():
    show_flash()
    st.title("📊 Rekap Nilai")
    quizzes = list_quizzes()
    if not quizzes:
        st.info("Belum ada kuis.")
        return
    ids = [q["id"] for q in quizzes]
    labels = {q["id"]: quiz_label(q) for q in quizzes}
    if st.session_state.get("rekap_quiz") not in ids:
        act = st.session_state.get("active_quiz")
        st.session_state["rekap_quiz"] = act if act in ids else ids[0]
    qid = st.selectbox("Kuis", ids, key="rekap_quiz", format_func=lambda i: labels[i])
    df = results_df(qid)
    if df.empty:
        st.info("Belum ada LJK yang disimpan untuk kuis ini.")
        st.button("📸 Mulai scan", key="rekap_to_scan", type="primary", on_click=goto, args=(NAV_SCAN, qid))
        return

    kkm = st.number_input("KKM (persen)", min_value=0, max_value=100, value=70, step=5, key="kkm")
    kelas_opts = [ALL_CLASSES] + sorted(df["Kelas"].unique(), key=str.lower)
    kelas = st.selectbox("Kelas", kelas_opts, key="rekap_kelas") if len(kelas_opts) > 2 else ALL_CLASSES
    view = df if kelas == ALL_CLASSES else df[df["Kelas"] == kelas]
    s = score_stats(view, kkm)
    metric_grid([("Peserta", s["peserta"]), ("Rata-rata", s["rata"]),
                 ("Tertinggi", s["maks"], "ok"), ("Terendah", s["min"], "bad"),
                 (f"Lulus (≥{kkm})", s["lulus"], "ok"), ("Belum lulus", s["peserta"] - s["lulus"], "warn")])

    st.dataframe(view.drop(columns=["Waktu Scan"]), hide_index=True)

    data, rows = export_bytes(qid, "xlsx")
    st.download_button("⬇️ Ekspor nilai ke Excel", data,
                       file_name=f"nilai_{_safe_name(labels[qid])}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       key="dl_xlsx", type="primary")
    csv_bytes, _ = export_bytes(qid, "csv")
    st.download_button("⬇️ Ekspor CSV", csv_bytes, file_name=f"nilai_{_safe_name(labels[qid])}.csv",
                       mime="text/csv", key="dl_csv")

    with st.expander("📈 Sebaran nilai"):
        st.bar_chart(score_distribution(view))
    if len(kelas_opts) > 2:
        with st.expander("🏫 Ringkasan per kelas"):
            st.dataframe(class_summary(df), hide_index=True)
    with st.expander("🔎 Analisis butir soal"):
        st.dataframe(item_analysis(qid), hide_index=True)
    with st.expander("⏳ Siswa belum di-scan"):
        students = list_students()
        pending = students[~students["NIS"].isin(scanned_nis(qid))]
        if pending.empty:
            st.success("Semua siswa sudah di-scan.")
        else:
            st.dataframe(pending, hide_index=True)


# ==========================================================================
# MAIN
# ==========================================================================
def login_gate():
    """True bila boleh lanjut. Bila APP_PASSWORD diset, minta kata sandi dulu."""
    expected = configured_password()
    if not expected or st.session_state.get("authed"):
        return True
    st.markdown('<div class="hero"><div class="t">🔒 Pemeriksa LJK</div>'
                '<div class="s">Masukkan kata sandi untuk melanjutkan</div></div>', unsafe_allow_html=True)
    with st.container(border=True):
        pw = st.text_input("Kata sandi", type="password", key="pw_input")
        if st.button("Masuk", type="primary", key="pw_btn"):
            if check_password(pw, expected):
                st.session_state["authed"] = True
                st.rerun()
            else:
                st.error("Kata sandi salah.")
    return False


PAGES = {NAV_KUIS: page_kuis, NAV_SCAN: page_scan, NAV_SISWA: page_siswa, NAV_REKAP: page_rekap}


def main():
    st.set_page_config(page_title="Pemeriksa LJK", page_icon="📝", layout="centered",
                       initial_sidebar_state="collapsed")
    inject_css()
    inject_pwa()
    if not login_gate():
        return
    dm.init_db()

    if st.session_state.get("nav") not in NAV_ITEMS:
        st.session_state["nav"] = NAV_KUIS
    with st.container(key="bottom_nav"):
        page = st.radio("Menu", NAV_ITEMS, key="nav", horizontal=True, label_visibility="collapsed")
    PAGES[page]()


if __name__ == "__main__":
    main()
