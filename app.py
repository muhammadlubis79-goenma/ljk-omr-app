#!/usr/bin/env python3
"""Pemeriksa LJK OMR - Desain Simple & Kontras Tinggi (gaya ZipGrade)."""

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
CHOICES = list(dm.VALID_OPTIONS)

MODE_100 = "Sama rata, skor maksimum 100"
MODE_1 = "Sama rata, 1 poin per soal"
MODE_MANUAL = "Atur manual per soal"
WEIGHT_MODES = [MODE_100, MODE_1, MODE_MANUAL]

ALL_CLASSES = "(Semua kelas)"
DEFAULT_FILL_THRESHOLD = 0.40
LOW_RES_SIDE = 720
MAX_IMAGE_SIDE = 3200

NAV_KUIS = "Kuis"
NAV_SCAN = "Scan"
NAV_SISWA = "Siswa"
NAV_REKAP = "Rekap"
NAV_ITEMS = [NAV_KUIS, NAV_SCAN, NAV_SISWA, NAV_REKAP]

SRC_CAMERA = "Kamera Langsung"
SRC_UPLOAD = "Unggah / Kamera HP"


# ==========================================================================
# LAPISAN LOGIKA
# ==========================================================================
def _rows(sql, params=()):
    with dm.get_connection() as conn:
        return [dict(r) for r in conn.execute(sql, params).fetchall()]


def _safe_name(text, fallback="x"):
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "_", str(text)).strip("_")
    return cleaned or fallback


def list_quizzes():
    return _rows(
        "SELECT q.id, q.quiz_name AS name, q.total_questions AS total_questions,"
        "       (SELECT COUNT(DISTINCT r.student_id) FROM scan_results r"
        "        WHERE r.quiz_id = q.id) AS peserta "
        "FROM quizzes q ORDER BY q.id DESC"
    )


def get_quiz(quiz_id):
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
    if not order:
        return None
    start = order.index(after) + 1 if after in order else 0
    for step in range(len(order)):
        nis = order[(start + step) % len(order)]
        if nis not in scanned:
            return nis
    return after if after in order else order[0]


def parse_key_string(text, n):
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


def make_pdf_bytes(quiz_name, total_questions, school, quiz_id, tanggal_iso):
    with tempfile.TemporaryDirectory() as tmp:
        path = os.path.join(tmp, "ljk.pdf")
        gen.generate_ljk_pdf(
            quiz_name, total_questions, path,
            nama_sekolah=school, quiz_id=str(quiz_id), tanggal=tanggal_iso,
        )
        with open(path, "rb") as fh:
            return fh.read()


def import_students_bytes(data, filename):
    ext = os.path.splitext(filename or "")[1].lower() or ".xlsx"
    with tempfile.TemporaryDirectory() as tmp:
        path = os.path.join(tmp, "siswa" + ext)
        with open(path, "wb") as fh:
            fh.write(data)
        return dm.import_students_from_excel(path)


def student_template_bytes():
    from openpyxl import Workbook
    from openpyxl.styles import Font

    wb = Workbook()
    ws = wb.active
    ws.title = "Siswa"
    ws.append(["NIS", "Nama_Lengkap", "Kelas"])
    ws.append(["0001", "Contoh Siswa Satu", "8A"])
    ws.append(["0002", "Contoh Siswa Dua", "8A"])
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


def details_df(grade):
    rows = []
    for d in grade["details"]:
        ans = d["jawaban"]
        if ans is None:
            status = "Kosong"
        elif len(ans) > 1:
            status = "Ganda"
        elif d["benar"]:
            status = "Benar"
        else:
            status = "Salah"
        rows.append({
            "No": d["nomor"],
            "Jawaban": ans or "-",
            "Kunci": d["kunci"],
            "Status": status,
            "Bobot": round(d["bobot"], 2),
        })
    return pd.DataFrame(rows)


def run_scan(image_bytes, quiz, fill_threshold=DEFAULT_FILL_THRESHOLD, show_key=False):
    img = cv2.imdecode(np.frombuffer(image_bytes, np.uint8), cv2.IMREAD_COLOR)
    if img is None:
        raise omr.OMRError("File bukan gambar yang bisa dibaca. Gunakan foto JPG / PNG.")

    h, w = img.shape[:2]
    notes = []
    if min(h, w) < LOW_RES_SIDE:
        notes.append(
            f"Foto beresolusi rendah ({w}x{h}). Gunakan 'Unggah / Kamera HP' agar memakai kamera bawaan."
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
                f"QR pada LJK menunjuk kuis ID '{meta.get('Quiz_ID')}', sedangkan kuis aktif "
                f"ber-ID '{quiz['id']}'."
            )
    else:
        notes.append("QR code tidak terbaca, kecocokan kuis tidak diperiksa.")

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
    with tempfile.TemporaryDirectory() as tmp:
        path = os.path.join(tmp, f"nilai.{fmt}")
        rows = dm.export_grades_to_excel(quiz_id, path)
        with open(path, "rb") as fh:
            return fh.read(), rows


def check_password(entered, expected):
    return hmac.compare_digest(
        hashlib.sha256(str(entered).encode("utf-8")).digest(),
        hashlib.sha256(str(expected).encode("utf-8")).digest(),
    )


def configured_password():
    value = None
    try:
        value = st.secrets.get("APP_PASSWORD")
    except Exception:
        value = None
    value = value or os.environ.get("APP_PASSWORD")
    return str(value) if value else None


def app_icon_png_b64(size=192):
    img = np.zeros((size, size, 3), np.uint8)
    img[:] = (50, 125, 46)
    c = size // 2
    cv2.circle(img, (c, c), int(size * 0.30), (255, 255, 255), max(4, size // 24), cv2.LINE_AA)
    cv2.circle(img, (c, c), int(size * 0.13), (255, 255, 255), -1, cv2.LINE_AA)
    ok, buf = cv2.imencode(".png", img)
    return base64.b64encode(buf.tobytes()).decode("ascii") if ok else ""


# ==========================================================================
# TAMPILAN: CSS SIMPLE GAYA ZIPGRADE & ANTI-DARK-MODE CONFLICT
# ==========================================================================
CSS = """
<style>
:root{
  --brand:#2e7d32; --brand-dark:#1b5e20; --brand-soft:#e8f5e9;
  --ok:#15803d; --bad:#b91c1c; --warn:#b45309;
  --bg:#f4f6f4; --card:#ffffff; --ink:#111827; --muted:#4b5563; --line:#d1d5db;
  --text-color:#111827; --background-color:#f4f6f4; --secondary-background-color:#ffffff;
}

/* Paksa tema terang & teks gelap agar HP Mode Gelap tidak membuat tulisan putih di atas putih */
html, body, [data-testid="stApp"], [data-testid="stAppViewContainer"], [data-testid="stMain"]{
  background-color:var(--bg) !important;
  color:var(--ink) !important;
  color-scheme:light !important;
}
header[data-testid="stHeader"], [data-testid="stToolbar"], [data-testid="stDecoration"],
[data-testid="stStatusWidget"], .stAppDeployButton, #MainMenu, footer{ display:none !important; }

/* Kontainer utama rapat & bersih */
.block-container{
  max-width:540px !important;
  padding:0 0.85rem 5.5rem !important;
}

/* Pastikan semua label, judul, paragraf, dan caption terbaca jelas */
h1, h2, h3, p, span, label, summary,
div[data-testid="stMarkdownContainer"] p,
[data-testid="stWidgetLabel"] p,
[data-testid="stCaptionContainer"] p,
[data-testid="stExpander"] summary p{
  color:var(--ink) !important;
}
[data-testid="stCaptionContainer"] p{
  color:var(--muted) !important;
  font-size:0.85rem !important;
}

/* Top Bar Hijau Flat ala ZipGrade */
.topbar{
  background:var(--brand-dark);
  color:#ffffff !important;
  margin:0 -0.85rem 0.85rem;
  padding:14px 16px;
  font-size:1.05rem;
  font-weight:800;
  letter-spacing:0.04em;
  text-transform:uppercase;
}

/* Kartu & Daftar Datar */
div[data-testid="stVerticalBlockBorderWrapper"]{
  background:var(--card) !important;
  border:1px solid var(--line) !important;
  border-radius:8px !important;
  box-shadow:none !important;
  padding:0.5rem 0.65rem !important;
  margin-bottom:0.5rem !important;
}

/* Input & Dropdown terang dengan teks hitam tajam */
input, textarea,
[data-baseweb="input"], [data-baseweb="input"] input,
[data-baseweb="textarea"], [data-baseweb="textarea"] textarea,
[data-baseweb="select"] > div{
  background-color:#ffffff !important;
  color:#111827 !important;
  -webkit-text-fill-color:#111827 !important;
  border-color:#cbd5e1 !important;
  border-radius:6px !important;
  font-size:16px !important;
}
input::placeholder, textarea::placeholder{
  color:#6b7280 !important;
  -webkit-text-fill-color:#6b7280 !important;
}
[data-testid="stNumberInput"] button{
  background-color:#f3f4f6 !important;
  color:#111827 !important;
}
[data-testid="stExpander"]{
  background-color:#ffffff !important;
  border:1px solid var(--line) !important;
  border-radius:8px !important;
}

/* Grid Angka Ringkas */
.mgrid{ display:grid; grid-template-columns:repeat(auto-fit,minmax(88px,1fr)); gap:6px; margin:6px 0 10px; }
.mcell{ background:#ffffff; border:1px solid var(--line); border-radius:6px; padding:8px 4px; text-align:center; }
.mcell .v{ font-size:1.25rem; font-weight:800; color:var(--ink); line-height:1.15; }
.mcell .l{ font-size:0.72rem; color:var(--muted); margin-top:2px; }
.mcell.ok .v{ color:var(--ok); } .mcell.bad .v{ color:var(--bad); } .mcell.warn .v{ color:var(--warn); }

.score{ text-align:center; padding:4px 0; }
.score .big{ font-size:2.4rem; font-weight:900; color:var(--brand-dark); line-height:1; }
.score .sub{ color:var(--muted); font-size:0.9rem; margin-top:4px; }

.qrow{ display:flex; justify-content:space-between; align-items:flex-start; gap:8px; }
.who{ font-weight:700; font-size:0.98rem; color:var(--ink); }
.meta{ color:var(--muted); font-size:0.82rem; margin-top:2px; }
.qright{ text-align:right; white-space:nowrap; font-size:0.82rem; color:var(--muted); }
.pill{ display:inline-block; padding:1px 7px; border-radius:4px; font-size:0.72rem; font-weight:700;
  background:var(--brand-soft); color:var(--brand-dark); margin-left:4px; }

/* Tombol Flat & Tegas */
.stButton > button, .stDownloadButton > button, [data-testid="stFormSubmitButton"] > button{
  width:100%;
  min-height:2.85rem;
  border-radius:6px !important;
  font-weight:700 !important;
  font-size:0.95rem !important;
  background-color:#ffffff !important;
  color:#111827 !important;
  border:1px solid #cbd5e1 !important;
  box-shadow:none !important;
}
.stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"]{
  background-color:var(--brand) !important;
  border-color:var(--brand) !important;
  color:#ffffff !important;
}
.stButton > button[kind="primary"] p, .stDownloadButton > button[kind="primary"] p{
  color:#ffffff !important;
}
.stButton > button[kind="primary"]:hover{
  background-color:var(--brand-dark) !important;
  border-color:var(--brand-dark) !important;
}

/* Kamera */
[data-testid="stCameraInput"]{ width:100%; }
[data-testid="stCameraInput"] video, [data-testid="stCameraInput"] img{
  width:100% !important; height:auto !important; border-radius:6px; display:block;
}
[data-testid="stCameraInputButton"]{
  width:100% !important; margin:8px 0 2px !important; min-height:3.6rem !important;
  border-radius:8px !important; font-size:1.05rem !important; font-weight:800 !important;
  background:var(--brand) !important; color:#ffffff !important; border:0 !important;
}
[data-testid="stImage"] img{ width:100% !important; border-radius:6px; }

/* Bottom Navigation Bar: Bersih tanpa bulatan radio & tidak terpotong di ujung kanan */
.st-key-bottom_nav{
  position:fixed; left:50%; bottom:0; transform:translateX(-50%);
  width:min(100%,540px); z-index:1000;
  background:#ffffff !important;
  border-top:1px solid var(--line);
  padding:4px 48px calc(4px + env(safe-area-inset-bottom)) 6px;
}
.st-key-bottom_nav [role="radiogroup"]{
  display:flex !important;
  flex-direction:row !important;
  flex-wrap:nowrap !important;
  gap:4px !important;
  width:100% !important;
}
.st-key-bottom_nav label{
  flex:1 1 0 !important;
  min-width:0 !important;
  margin:0 !important;
  padding:10px 4px !important;
  border-radius:6px !important;
  cursor:pointer;
  display:flex !important;
  justify-content:center !important;
  align-items:center !important;
  text-align:center !important;
  min-height:2.75rem !important;
  background:transparent !important;
}
/* Sembunyikan bulatan radio hitam pada semua versi Streamlit */
.st-key-bottom_nav label > div:first-of-type,
.st-key-bottom_nav label input[type="radio"]{
  display:none !important;
  width:0 !important;
  height:0 !important;
  margin:0 !important;
  padding:0 !important;
}
.st-key-bottom_nav label div[data-testid="stMarkdownContainer"] p{
  font-size:0.88rem !important;
  font-weight:700 !important;
  white-space:nowrap !important;
  margin:0 !important;
  color:var(--muted) !important;
}
.st-key-bottom_nav label:has(input:checked){
  background:var(--brand-soft) !important;
  border-bottom:3px solid var(--brand-dark) !important;
}
.st-key-bottom_nav label:has(input:checked) p{
  color:var(--brand-dark) !important;
}
.st-key-bottom_nav [data-testid="stWidgetLabel"]{ display:none !important; }
</style>
"""


def inject_css():
    st.markdown(CSS, unsafe_allow_html=True)


def inject_pwa():
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
  meta('theme-color', '#1b5e20');
  meta('mobile-web-app-capable', 'yes');
  meta('apple-mobile-web-app-capable', 'yes');
  meta('apple-mobile-web-app-title', 'LJK');
  const icon = __ICON__;
  const link = (rel, href) => {
    let l = h.querySelector('link[rel="' + rel + '"]');
    if (!l) { l = d.createElement('link'); l.setAttribute('rel', rel); h.appendChild(l); }
    l.setAttribute('href', href);
  };
  link('apple-touch-icon', icon);
  const manifest = {name: 'Pemeriksa LJK', short_name: 'LJK', display: 'standalone',
    background_color: '#f4f6f4', theme_color: '#1b5e20',
    start_url: window.parent.location.pathname || '/',
    icons: [{src: icon, sizes: '192x192', type: 'image/png'}]};
  link('manifest', 'data:application/manifest+json,' + encodeURIComponent(JSON.stringify(manifest)));
} catch (e) {}
</script>
""".replace("__ICON__", json.dumps(icon))
        components.html(script, height=0)
    except Exception:
        pass


def esc(text):
    return html.escape(str(text))


def topbar(title):
    st.markdown(f'<div class="topbar">{esc(title)}</div>', unsafe_allow_html=True)


def metric_grid(items):
    cells = "".join(
        f'<div class="mcell {c[2] if len(c) > 2 else ""}"><div class="v">{esc(c[1])}</div>'
        f'<div class="l">{esc(c[0])}</div></div>'
        for c in items
    )
    st.markdown(f'<div class="mgrid">{cells}</div>', unsafe_allow_html=True)


def flash(level, text):
    st.session_state["flash"] = (level, text)


def show_flash():
    item = st.session_state.pop("flash", None)
    if not item:
        return
    level, text = item
    {"success": st.success, "error": st.error, "warning": st.warning}.get(level, st.info)(text)


def goto(page, quiz_id=None):
    st.session_state["nav"] = page
    if quiz_id is not None:
        st.session_state["active_quiz"] = quiz_id
        st.session_state["scan_quiz"] = quiz_id


@st.cache_data(show_spinner=False, max_entries=20)
def cached_pdf(quiz_name, total_questions, school, quiz_id, tanggal_iso):
    return make_pdf_bytes(quiz_name, total_questions, school, quiz_id, tanggal_iso)


def quiz_label(q):
    return f"{q['name']} ({q['total_questions']} soal)"


# ==========================================================================
# HALAMAN KUIS (Simple ala ZipGrade)
# ==========================================================================
def page_kuis():
    if st.session_state.pop("reset_form", False):
        st.session_state["new_name"] = ""
        st.session_state["quick_key"] = ""
        st.session_state["w_mode"] = MODE_100
        st.session_state["show_new_quiz"] = False

    topbar("Kuis")
    show_flash()

    # Tombol buka/tutup form Kuis Baru (agar halaman utama tetap bersih seperti ZipGrade)
    is_creating = st.session_state.get("show_new_quiz", False)
    if not is_creating:
        if st.button("+ Kuis Baru", type="primary", key="open_new_quiz"):
            st.session_state["show_new_quiz"] = True
            st.rerun()
    else:
        with st.container(border=True):
            st.subheader("Buat Kuis Baru")
            name = st.text_input("Nama kuis", key="new_name", placeholder="mis. Informatika SAS Ganjil")
            n = int(st.number_input("Jumlah soal", min_value=MIN_SOAL, max_value=MAX_SOAL,
                                    value=30, step=1, key="new_n"))
            quick = st.text_area(
                "Kunci Jawaban",
                key="quick_key", height=95,
                placeholder="Ketik/tempel huruf A-E, mis. ABCDEABCDE...",
            )
            letters = [c for c in quick.upper() if c.isalpha()]
            invalid = sorted({c for c in letters if c not in CHOICES})
            if not letters:
                st.caption(f"Butuh {n} huruf (A-E).")
            elif invalid:
                st.error(f"Huruf tidak valid: {', '.join(invalid)}. Gunakan A-E saja.")
            elif len(letters) == n:
                st.success(f"Kunci lengkap ({n}/{n} soal).")
            else:
                st.warning(f"Terisi {len(letters)} dari {n} soal.")

            with st.expander("Atur bobot nilai (opsional)"):
                mode = st.radio("Mode bobot", WEIGHT_MODES, key="w_mode")
                manual_values = None
                if mode == MODE_MANUAL:
                    w_df = pd.DataFrame({"No": list(range(1, n + 1)), "Bobot": [1.0] * n})
                    w_edit = st.data_editor(
                        w_df, key=f"w_editor_{n}", hide_index=True, disabled=["No"],
                        num_rows="fixed", height=280,
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
                    st.caption(f"Skor maksimum: {sum(weights):g}")

            c1, c2 = st.columns(2)
            with c1:
                if st.button("Simpan Kuis", type="primary", key="save_quiz"):
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
                        flash("success", f"Kuis '{name.strip()}' berhasil dibuat.")
                        st.rerun()
            with c2:
                if st.button("Batal", key="cancel_new_quiz"):
                    st.session_state["show_new_quiz"] = False
                    st.rerun()

    # Baris Sortir & Cari ala ZipGrade
    quizzes = list_quizzes()
    if not quizzes:
        st.info("Belum ada kuis. Tekan tombol '+ Kuis Baru' di atas untuk memulai.")
    else:
        col_sort, col_search = st.columns([1, 1.4])
        with col_sort:
            sort_by = st.selectbox("Sortir", ["Terbaru", "Nama A-Z"], key="quiz_sort")
        with col_search:
            search_q = st.text_input("Cari", key="quiz_search", placeholder="Cari nama kuis...")

        filtered = [
            q for q in quizzes
            if not search_q.strip() or search_q.strip().lower() in q["name"].lower()
        ]
        if sort_by == "Nama A-Z":
            filtered.sort(key=lambda item: item["name"].lower())

        school = st.session_state.get("school_name") or "Nama Sekolah"
        for q in filtered:
            with st.container(border=True):
                active = st.session_state.get("active_quiz") == q["id"]
                badge = '<span class="pill">Aktif</span>' if active else ""
                st.markdown(
                    f'<div class="qrow">'
                    f'  <div><div class="who">{esc(q["name"])}{badge}</div>'
                    f'  <div class="meta">{q["total_questions"]} soal · ID {q["id"]}</div></div>'
                    f'  <div class="qright">Kertas: <b>{q["peserta"]}</b></div>'
                    f'</div>',
                    unsafe_allow_html=True,
                )
                b1, b2, b3 = st.columns([1.4, 1.2, 0.7])
                with b1:
                    st.button("Scan", key=f"scan_{q['id']}", type="primary",
                              on_click=goto, args=(NAV_SCAN, q["id"]))
                with b2:
                    if st.session_state.get("pdf_for") == q["id"]:
                        pdf = cached_pdf(q["name"], q["total_questions"], school, q["id"],
                                         date.today().isoformat())
                        st.download_button("Unduh PDF", pdf,
                                           file_name=f"LJK_{_safe_name(q['name'])}.pdf",
                                           mime="application/pdf", key=f"dl_pdf_{q['id']}")
                    else:
                        if st.button("Cetak LJK", key=f"mk_pdf_{q['id']}"):
                            st.session_state["pdf_for"] = q["id"]
                            st.rerun()
                with b3:
                    if st.button("Hapus", key=f"del_{q['id']}"):
                        st.session_state["del_quiz"] = q["id"]
                        st.rerun()

                if st.session_state.get("del_quiz") == q["id"]:
                    st.warning("Hapus kuis ini beserta semua nilainya?")
                    dy, dn = st.columns(2)
                    with dy:
                        if st.button("Ya, Hapus", key=f"del_yes_{q['id']}", type="primary"):
                            dm.delete_quiz(q["id"])
                            st.session_state.pop("del_quiz", None)
                            st.session_state.pop("pdf_for", None)
                            if st.session_state.get("active_quiz") == q["id"]:
                                st.session_state.pop("active_quiz", None)
                            flash("success", f"Kuis '{q['name']}' dihapus.")
                            st.rerun()
                    with dn:
                        if st.button("Batal", key=f"del_no_{q['id']}"):
                            st.session_state.pop("del_quiz", None)
                            st.rerun()

    # Pengaturan Tambahan di Bawah
    with st.expander("Pengaturan LJK & Cadangan Data"):
        st.text_input("Nama sekolah (tercetak di header LJK)", key="school_name", value="Nama Sekolah")
        st.download_button("Unduh Cadangan Database (.db)", dm.backup_bytes(),
                           file_name=f"cadangan_ljk_{datetime.now():%Y%m%d_%H%M}.db",
                           mime="application/octet-stream", key="dl_backup")
        up = st.file_uploader("Pulihkan dari cadangan (.db)", type=["db", "sqlite"], key="restore_up")
        agree = st.checkbox("Ganti data sekarang dengan file cadangan", key="restore_ok")
        if st.button("Pulihkan Data", key="restore_btn", disabled=not (up and agree)):
            try:
                dm.restore_from_bytes(up.getvalue())
            except ValueError as exc:
                st.error(str(exc))
            else:
                st.session_state.pop("active_quiz", None)
                flash("success", "Data berhasil dipulihkan.")
                st.rerun()


# ==========================================================================
# HALAMAN SCAN
# ==========================================================================
def _cb_bump_camera():
    st.session_state["cam_ver"] = st.session_state.get("cam_ver", 0) + 1


def _cb_undo(scan_id, nis):
    if undo_scan(scan_id):
        flash("success", "Penyimpanan terakhir dibatalkan.")
        st.session_state["scan_stu"] = nis
    st.session_state.pop("last_saved", None)


def page_scan():
    topbar("Scan LJK")
    show_flash()
    quizzes = list_quizzes()
    if not quizzes:
        st.info("Belum ada kuis. Buat kuis dulu di menu Kuis.")
        st.button("Ke Menu Kuis", key="to_kuis", type="primary", on_click=goto, args=(NAV_KUIS,))
        return

    ids = [q["id"] for q in quizzes]
    labels = {q["id"]: quiz_label(q) for q in quizzes}
    if st.session_state.get("scan_quiz") not in ids:
        act = st.session_state.get("active_quiz")
        st.session_state["scan_quiz"] = act if act in ids else ids[0]
    qid = st.selectbox("Pilih Kuis", ids, key="scan_quiz", format_func=lambda i: labels[i])
    st.session_state["active_quiz"] = qid
    quiz = get_quiz(qid)

    students = list_students()
    if students.empty:
        st.warning("Belum ada data siswa. Tambahkan siswa dulu di menu Siswa.")
        st.button("Ke Menu Siswa", key="to_siswa", type="primary", on_click=goto, args=(NAV_SISWA,))
        return

    classes = sorted(students["Kelas"].unique(), key=str.lower)
    if len(classes) > 1:
        if st.session_state.get("scan_kelas") not in [ALL_CLASSES] + classes:
            st.session_state["scan_kelas"] = ALL_CLASSES
        kelas = st.selectbox("Filter Kelas", [ALL_CLASSES] + classes, key="scan_kelas")
        subset = students if kelas == ALL_CLASSES else students[students["Kelas"] == kelas]
    else:
        subset = students

    order = list(subset["NIS"])
    info = {r.NIS: (r.Nama, r.Kelas) for r in students.itertuples(index=False)}
    scanned = scanned_nis(qid)

    advance = st.session_state.pop("advance_after", None)
    if advance is not None or st.session_state.get("scan_stu") not in order:
        st.session_state["scan_stu"] = next_unscanned(order, scanned, advance)

    def stu_label(nis):
        nama, kls = info[nis]
        status = "[Sudah] " if nis in scanned else ""
        return f"{status}{nama} ({kls})"

    nis = st.selectbox("Pilih Siswa", order, key="scan_stu", format_func=stu_label)
    nama, kls = info[nis]
    done = len(scanned & set(order))
    st.caption(f"Progres: {done} dari {len(order)} siswa sudah di-scan.")

    with st.expander("Opsi Scan"):
        auto_save = st.checkbox("Simpan otomatis setelah terbaca", key="auto_save")
        threshold = st.slider("Sensitivitas arsiran", 0.20, 0.70, DEFAULT_FILL_THRESHOLD, 0.05, key="fill_thr")
        show_key = st.checkbox("Tampilkan kunci benar pada soal salah", key="show_key")

    src = st.radio("Metode Foto", [SRC_UPLOAD, SRC_CAMERA], key="scan_src", horizontal=True)
    ver = st.session_state.setdefault("cam_ver", 0)
    if src == SRC_CAMERA:
        photo = st.camera_input("Jepret LJK", key=f"cam_{ver}", label_visibility="collapsed")
    else:
        photo = st.file_uploader("Ambil / Pilih Foto LJK", type=["jpg", "jpeg", "png"],
                                 key=f"up_{ver}", label_visibility="collapsed")
        st.caption("Disarankan: tombol ini membuka kamera asli HP beresolusi tinggi.")

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
        with st.spinner("Memeriksa LJK..."):
            try:
                scan = run_scan(data, quiz, threshold, show_key)
            except (omr.OMRError, ValueError) as exc:
                error = str(exc)
        st.session_state["scan_cache"] = {"sig": sig, "scan": scan, "error": error}

    if error:
        with st.container(border=True):
            st.error(error)
            st.button("Foto Ulang", key="retry_err", type="primary", on_click=_cb_bump_camera)
        return

    grade = scan["grade"]
    with st.container(border=True):
        st.markdown(f'<div class="who">{esc(nama)}</div><div class="meta">{esc(kls)} · NIS {esc(nis)}</div>',
                    unsafe_allow_html=True)
        st.markdown(
            f'<div class="score"><div class="big">{grade["total_score"]:g}</div>'
            f'<div class="sub">dari {grade["max_score"]:g} ({grade["persentase"]:g}%)</div></div>',
            unsafe_allow_html=True,
        )
        salah_murni = max(grade["salah"] - grade["kosong"] - grade["ganda"], 0)
        metric_grid([("Benar", grade["benar"], "ok"), ("Salah", salah_murni, "bad"),
                     ("Kosong", grade["kosong"]), ("Ganda", grade["ganda"], "warn")])
        if scan["mismatch"]:
            st.warning(scan["mismatch"])
        for note in scan["notes"]:
            st.warning(note)
        st.image(cv2.cvtColor(resize_to_width(scan["feedback_bgr"], 1000), cv2.COLOR_BGR2RGB),
                 caption="Hijau = benar · Merah = salah")
        with st.expander("Rincian per nomor soal"):
            st.dataframe(details_df(grade), hide_index=True)

        blocked = bool(scan["mismatch"])
        c1, c2 = st.columns(2)
        with c1:
            if st.button("Simpan Nilai", key="save_next", type="primary"):
                _do_save(scan, quiz, nis, nama, kls, data, photo_name)
        with c2:
            st.button("Foto Ulang", key="retry_ok", on_click=_cb_bump_camera)

    if auto_save and not blocked and st.session_state.get("auto_done") != sig:
        st.session_state["auto_done"] = sig
        _do_save(scan, quiz, nis, nama, kls, data, photo_name)


def _do_save(scan, quiz, nis, nama, kls, data, photo_name):
    try:
        sid = save_scan(scan, quiz, nis, data, photo_name)
    except Exception as exc:
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
    flash("success", f"Tersimpan: {nama} ({grade['total_score']:g}/{grade['max_score']:g})")
    st.rerun()


def _render_last_saved(last):
    with st.container(border=True):
        st.markdown(
            f'<div class="meta">Terakhir disimpan</div><div class="who">{esc(last["nama"])} · '
            f'{last["skor"]:g}/{last["maks"]:g} ({last["persen"]:g}%)</div>',
            unsafe_allow_html=True,
        )
        st.button("Batalkan Simpan Terakhir", key="undo_last", on_click=_cb_undo,
                  args=(last["id"], last["nis"]))


# ==========================================================================
# HALAMAN SISWA
# ==========================================================================
def page_siswa():
    if st.session_state.pop("reset_add", False):
        st.session_state["add_nis"] = ""
        st.session_state["add_nama"] = ""
    topbar("Data Siswa")
    show_flash()
    students = list_students()
    metric_grid([("Total Siswa", len(students)), ("Jumlah Kelas", students["Kelas"].nunique() if len(students) else 0)])

    with st.expander("Impor Siswa dari Excel / CSV"):
        st.caption("Format kolom wajib: NIS, Nama_Lengkap, Kelas")
        st.download_button("Unduh Template Excel", student_template_bytes(),
                           file_name="template_siswa.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           key="dl_tpl")
        up = st.file_uploader("Pilih file siswa", type=["xlsx", "csv"], key="stu_up")
        if st.button("Impor Sekarang", key="stu_import", type="primary", disabled=up is None):
            try:
                res = import_students_bytes(up.getvalue(), up.name)
            except Exception as exc:
                st.error(f"Gagal impor: {exc}")
            else:
                msg = f"Baru: {res['inserted']} · Diperbarui: {res['updated']} · Dilewati: {res['skipped']}"
                flash("success" if not res.get("errors") else "warning", msg)
                st.rerun()

    with st.expander("Tambah Satu Siswa Manual"):
        c_nis = st.text_input("NIS", key="add_nis")
        c_nama = st.text_input("Nama lengkap", key="add_nama")
        c_kelas = st.text_input("Kelas", key="add_kelas", placeholder="mis. 8A")
        if st.button("Simpan Siswa", key="add_btn", type="primary"):
            try:
                status = dm.add_student(c_nis, c_nama, c_kelas)
            except ValueError as exc:
                st.error(str(exc))
            else:
                st.session_state["reset_add"] = True
                flash("success", f"Siswa {c_nama.strip()} ({status}).")
                st.rerun()

    if students.empty:
        st.info("Belum ada data siswa.")
        return
    classes = [ALL_CLASSES] + sorted(students["Kelas"].unique(), key=str.lower)
    kelas = st.selectbox("Filter Kelas", classes, key="stu_filter")
    view = students if kelas == ALL_CLASSES else students[students["Kelas"] == kelas]
    st.dataframe(view, hide_index=True, use_container_width=True)


# ==========================================================================
# HALAMAN REKAP
# ==========================================================================
def page_rekap():
    topbar("Rekap Nilai")
    show_flash()
    quizzes = list_quizzes()
    if not quizzes:
        st.info("Belum ada kuis.")
        return
    ids = [q["id"] for q in quizzes]
    labels = {q["id"]: quiz_label(q) for q in quizzes}
    if st.session_state.get("rekap_quiz") not in ids:
        act = st.session_state.get("active_quiz")
        st.session_state["rekap_quiz"] = act if act in ids else ids[0]
    qid = st.selectbox("Pilih Kuis", ids, key="rekap_quiz", format_func=lambda i: labels[i])
    df = results_df(qid)
    if df.empty:
        st.info("Belum ada LJK yang di-scan untuk kuis ini.")
        st.button("Mulai Scan", key="rekap_to_scan", type="primary", on_click=goto, args=(NAV_SCAN, qid))
        return

    c1, c2 = st.columns(2)
    with c1:
        kkm = st.number_input("KKM (%)", min_value=0, max_value=100, value=70, step=5, key="kkm")
    with c2:
        kelas_opts = [ALL_CLASSES] + sorted(df["Kelas"].unique(), key=str.lower)
        kelas = st.selectbox("Kelas", kelas_opts, key="rekap_kelas")

    view = df if kelas == ALL_CLASSES else df[df["Kelas"] == kelas]
    s = score_stats(view, kkm)
    metric_grid([("Peserta", s["peserta"]), ("Rata-rata", s["rata"]),
                 ("Tertinggi", s["maks"], "ok"), ("Terendah", s["min"], "bad"),
                 (f"Lulus (>={kkm})", s["lulus"], "ok")])

    st.dataframe(view.drop(columns=["Waktu Scan"]), hide_index=True, use_container_width=True)

    data, _ = export_bytes(qid, "xlsx")
    st.download_button("Unduh Nilai Excel (.xlsx)", data,
                       file_name=f"nilai_{_safe_name(labels[qid])}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       key="dl_xlsx", type="primary")

    with st.expander("Analisis Butir Soal"):
        st.dataframe(item_analysis(qid), hide_index=True, use_container_width=True)
    with st.expander("Siswa Belum Di-scan"):
        students = list_students()
        pending = students[~students["NIS"].isin(scanned_nis(qid))]
        if pending.empty:
            st.success("Semua siswa sudah di-scan.")
        else:
            st.dataframe(pending, hide_index=True, use_container_width=True)


# ==========================================================================
# MAIN
# ==========================================================================
def login_gate():
    expected = configured_password()
    if not expected or st.session_state.get("authed"):
        return True
    topbar("Masuk Aplikasi LJK")
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
    st.set_page_config(page_title="Pemeriksa LJK", page_icon="LJK", layout="centered",
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
