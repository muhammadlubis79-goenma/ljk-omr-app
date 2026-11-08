#!/usr/bin/env python3
"""Aplikasi LJK (Streamlit): cetak LJK, kelola siswa, scan & koreksi, laporan nilai.

Menggabungkan tiga modul:
    ljk_generator.py   -> membuat PDF LJK
    database_manager.py -> database SQLite + impor/ekspor Excel
    omr_engine.py      -> membaca foto LJK dan menilainya

Jalankan:
    pip install streamlit reportlab "qrcode[pil]" openpyxl opencv-python numpy pandas
    streamlit run app.py

Semua file (app.py, ljk_generator.py, database_manager.py, omr_engine.py) harus berada
di folder yang sama. Database `ljk_data.db` dan foto hasil scan (folder `scans/`) disimpan
di folder itu juga; ubah lokasinya dengan environment variable LJK_DB_PATH.
"""

import hashlib
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
QUESTION_OPTIONS = [15, 30, 45, 75, 100]
CHOICES = list(dm.VALID_OPTIONS)  # A-E, sama dengan tata letak LJK

MODE_100 = "Sama rata, skor maksimum 100"
MODE_1 = "Sama rata, 1 poin per soal"
MODE_MANUAL = "Atur manual per soal"
WEIGHT_MODES = [MODE_100, MODE_1, MODE_MANUAL]

ALL_CLASSES = "(Semua kelas)"
DEFAULT_FILL_THRESHOLD = 0.40


# ==========================================================================
# LAPISAN LOGIKA (tanpa streamlit, sehingga bisa diuji terpisah)
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


# ---- kunci jawaban & bobot -------------------------------------------------
def parse_key_string(text, n):
    """'ABCDE...' (boleh ada spasi / koma) -> list huruf sepanjang n."""
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

    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        res = omr.scan_ljk(
            img, quiz["key"], total_questions=quiz["total_questions"],
            weights=quiz["weights"], show_correct_key=show_key,
            fill_threshold=fill_threshold,
        )
    notes = [str(w.message) for w in caught if issubclass(w.category, RuntimeWarning)]

    meta = res["metadata"]
    mismatch = None
    if meta:
        if str(meta.get("Quiz_ID")) != str(quiz["id"]):
            mismatch = (
                f"QR code pada LJK ini menunjuk Quiz_ID '{meta.get('Quiz_ID')}', sedangkan kuis "
                f"yang dipilih memiliki ID '{quiz['id']}'. Kemungkinan LJK berasal dari kuis lain."
            )
    else:
        notes.append("QR code pada LJK tidak terbaca, jadi kecocokan LJK dengan kuis tidak bisa diperiksa.")

    return {
        "answers": res["student_answers"],
        "grade": res["grade"],
        "feedback_bgr": res["feedback"],
        "metadata": meta,
        "alignment_score": res["alignment_score"],
        "notes": notes,
        "mismatch": mismatch,
    }


def save_scan(scan, quiz, nis, photo_bytes, photo_name="foto.jpg", scans_dir=None):
    """Simpan foto asli + gambar koreksi ke disk lalu hasilnya ke database. Mengembalikan id scan."""
    if scans_dir is None:
        scans_dir = os.path.join(os.path.dirname(os.path.abspath(dm.DB_PATH)), "scans")
    folder = os.path.join(scans_dir, f"kuis_{quiz['id']}")
    os.makedirs(folder, exist_ok=True)

    ext = os.path.splitext(photo_name or "")[1].lower()
    if ext not in (".jpg", ".jpeg", ".png"):
        ext = ".jpg"
    base = f"{_safe_name(nis)}_{datetime.now():%Y%m%d_%H%M%S_%f}"
    orig_path = os.path.join(folder, base + ext)
    fb_path = os.path.join(folder, base + "_koreksi.png")

    written = []
    try:
        with open(orig_path, "wb") as fh:
            fh.write(photo_bytes)
        written.append(orig_path)
        ok, buf = cv2.imencode(".png", scan["feedback_bgr"])
        if ok:
            with open(fb_path, "wb") as fh:
                fh.write(buf.tobytes())
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
    out = (
        df.groupby("Kelas")
        .agg(Peserta=("NIS", "count"), Rata=("Persentase", "mean"),
             Tertinggi=("Persentase", "max"), Terendah=("Persentase", "min"))
        .round(2)
        .reset_index()
        .rename(columns={"Rata": "Rata-rata"})
    )
    return out


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


# ==========================================================================
# LAPISAN ANTARMUKA (Streamlit)
# ==========================================================================
@st.cache_data(show_spinner=False)
def cached_pdf(quiz_name, total_questions, school, quiz_id, tanggal_iso):
    return make_pdf_bytes(quiz_name, total_questions, school, quiz_id, tanggal_iso)


def flash(level, text):
    """Pesan yang tetap tampil setelah st.rerun()."""
    st.session_state["flash"] = (level, text)


def show_flash():
    item = st.session_state.get("flash")
    if not item:
        return
    del st.session_state["flash"]
    level, text = item
    {"success": st.success, "error": st.error, "warning": st.warning}.get(level, st.info)(text)


def _ensure_tables(n):
    """Siapkan tabel kunci/bobot di session_state; reset bila jumlah soal berubah."""
    if st.session_state.get("tables_n") != n:
        numbers = list(range(1, n + 1))
        st.session_state["key_df"] = pd.DataFrame({"No": numbers, "Kunci": [None] * n})
        st.session_state["w_df"] = pd.DataFrame({"No": numbers, "Bobot": [1.0] * n})
        st.session_state["tables_n"] = n
        st.session_state["editor_ver"] = st.session_state.get("editor_ver", 0) + 1


# ---- Halaman 1: Dashboard & Cetak LJK ------------------------------------------------
def page_dashboard():
    st.title("📊 Dashboard & Cetak LJK")
    # Form dikosongkan SETELAH simpan (di awal run berikutnya, sebelum widget dibuat, karena
    # Streamlit melarang mengubah nilai widget yang sudah dibuat pada run yang sama).
    if st.session_state.pop("reset_form", False):
        st.session_state["new_name"] = ""
        st.session_state["quick_key"] = ""
        st.session_state.pop("tables_n", None)
    counts = get_counts()
    cols = st.columns(4)
    cols[0].metric("Kelas", counts["kelas"])
    cols[1].metric("Siswa", counts["siswa"])
    cols[2].metric("Kuis", counts["kuis"])
    cols[3].metric("LJK ter-scan", counts["scan"])

    # ---- 1. Buat kuis ----
    st.divider()
    st.subheader("1. Buat kuis, kunci jawaban & bobot nilai")
    c_name, c_choice, c_n = st.columns([3, 2, 1])
    name = c_name.text_input("Nama kuis", key="new_name", placeholder="mis. Kuis Matematika Bab 3")
    c_choice.caption("Pilihan jawaban")
    c_choice.markdown("**A · B · C · D · E**")
    n = c_n.selectbox("Jumlah soal", QUESTION_OPTIONS, index=1, key="new_n")

    _ensure_tables(n)
    ver = st.session_state["editor_ver"]
    left, right = st.columns(2)

    with left:
        st.markdown("**Kunci jawaban**")
        quick = st.text_input("Isi cepat: tempel deretan huruf (mis. ABCDE…)", key="quick_key")
        if st.button("Terapkan ke tabel (menimpa semua kunci)", key="apply_quick"):
            try:
                letters = parse_key_string(quick, n)
            except ValueError as exc:
                st.error(str(exc))
            else:
                st.session_state["key_df"]["Kunci"] = letters
                st.session_state["editor_ver"] += 1
                st.rerun()
        key_edit = st.data_editor(
            st.session_state["key_df"],
            key=f"key_editor_{n}_{ver}",
            hide_index=True,
            disabled=["No"],
            num_rows="fixed",
            height=420,
            column_config={
                "No": st.column_config.NumberColumn("No"),
                "Kunci": st.column_config.SelectboxColumn("Kunci", options=CHOICES, required=False),
            },
        )

    weights = None
    with right:
        st.markdown("**Bobot nilai**")
        mode = st.radio("Mode bobot", WEIGHT_MODES, key="w_mode")
        manual_values = None
        if mode == MODE_MANUAL:
            w_edit = st.data_editor(
                st.session_state["w_df"],
                key=f"w_editor_{n}",
                hide_index=True,
                disabled=["No"],
                num_rows="fixed",
                height=420,
                column_config={
                    "No": st.column_config.NumberColumn("No"),
                    "Bobot": st.column_config.NumberColumn("Bobot", min_value=0.0, step=0.5, format="%.2f"),
                },
            )
            manual_values = [None if pd.isna(v) else float(v) for v in w_edit["Bobot"]]
        try:
            weights = build_weights(mode, n, manual_values)
        except ValueError as exc:
            st.error(str(exc))
        else:
            st.caption(f"Skor maksimum kuis: **{sum(weights):g}**")

    if st.button("💾 Simpan kuis", type="primary", key="save_quiz"):
        problems = []
        if not (name or "").strip():
            problems.append("Nama kuis wajib diisi.")
        keys = [None if pd.isna(v) else str(v) for v in key_edit["Kunci"]]
        missing = [i for i, k in enumerate(keys, start=1) if k not in CHOICES]
        if missing:
            shown = ", ".join(map(str, missing[:12])) + (" …" if len(missing) > 12 else "")
            problems.append(f"Kunci jawaban belum lengkap ({len(missing)} soal): nomor {shown}.")
        if weights is None:
            problems.append("Bobot nilai belum valid.")
        if problems:
            for msg in problems:
                st.error(msg)
        else:
            qid = dm.create_quiz(name.strip(), n, keys, weights)
            st.session_state["print_quiz"] = qid
            st.session_state["reset_form"] = True
            flash("success", f"Kuis '{name.strip()}' tersimpan (ID {qid}). LJK-nya siap diunduh di bagian 2.")
            st.rerun()

    # ---- 2. Cetak LJK ----
    st.divider()
    st.subheader("2. Cetak LJK")
    show_flash()
    quizzes = list_quizzes()
    if not quizzes:
        st.info("Belum ada kuis. Simpan kuis di bagian 1 untuk membuat LJK.")
        return

    ids = [q["id"] for q in quizzes]
    labels = {q["id"]: f"#{q['id']} · {q['name']} ({q['total_questions']} soal)" for q in quizzes}
    if st.session_state.get("print_quiz") not in ids:
        st.session_state["print_quiz"] = ids[0]
    st.session_state.setdefault("school_name", "Nama Sekolah")
    st.session_state.setdefault("print_date", date.today())

    qid = st.selectbox("Pilih kuis", ids, key="print_quiz", format_func=lambda i: labels[i])
    c1, c2 = st.columns(2)
    school = c1.text_input("Nama sekolah", key="school_name")
    tanggal = c2.date_input("Tanggal kuis", key="print_date")
    if not isinstance(tanggal, date):
        tanggal = date.today()

    quiz = get_quiz(qid)
    pdf = cached_pdf(quiz["name"], quiz["total_questions"], (school or "").strip() or "Nama Sekolah",
                     quiz["id"], tanggal.isoformat())
    st.download_button(
        "⬇️ Generate & Download PDF LJK",
        data=pdf,
        file_name=f"LJK_{_safe_name(quiz['name'], 'kuis')}_{quiz['total_questions']}soal.pdf",
        mime="application/pdf",
        type="primary",
        key="dl_pdf",
    )
    st.caption(
        f"QR code pada LJK berisi Quiz_ID = {quiz['id']}, jumlah soal, dan tanggal. Cetak dengan skala "
        "100% (Actual size), bukan 'Fit to page', agar kotak hitam sudut tetap 1×1 cm. "
        "Satu LJK untuk satu siswa; nama siswa ditulis tangan di kolom Nama."
    )

    with st.expander("Daftar kuis tersimpan"):
        st.dataframe(pd.DataFrame(quizzes).rename(columns={
            "id": "ID", "name": "Nama Kuis", "total_questions": "Jumlah Soal", "peserta": "Siswa ter-scan"}),
            hide_index=True)


# ---- Halaman 2: Kelola Data Siswa ------------------------------------------------------
def page_students():
    st.title("👥 Kelola Data Siswa")
    st.caption("Format file: kolom **NIS**, **Nama_Lengkap**, **Kelas**. Kelas baru dibuat otomatis; "
               "NIS yang sudah ada diperbarui, bukan digandakan.")

    c1, c2 = st.columns([3, 1])
    uploaded = c1.file_uploader("Upload file siswa (.xlsx / .csv)", type=["xlsx", "xlsm", "csv"],
                                key="students_upload")
    c2.download_button("📄 Unduh template", data=student_template_bytes(), file_name="template_siswa.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       key="dl_template")

    if uploaded is not None and st.button("📥 Import ke database", type="primary", key="do_import"):
        try:
            res = import_students_bytes(uploaded.getvalue(), uploaded.name)
        except (ValueError, FileNotFoundError, PermissionError) as exc:
            st.error(f"Import gagal: {exc}")
        except Exception as exc:  # file Excel rusak, dsb.
            st.error(f"File tidak bisa dibaca ({type(exc).__name__}). Pastikan file .xlsx / .csv yang valid.")
        else:
            m = st.columns(3)
            m[0].metric("Siswa baru", res["inserted"])
            m[1].metric("Diperbarui", res["updated"])
            m[2].metric("Dilewati", res["skipped"])
            if res["new_classes"]:
                st.success("Kelas baru dibuat: " + ", ".join(res["new_classes"]))
            if res["errors"]:
                with st.expander(f"Baris yang dilewati ({len(res['errors'])})", expanded=True):
                    for msg in res["errors"]:
                        st.write("• " + msg)

    st.divider()
    st.subheader("Daftar siswa")
    df = list_students()
    if df.empty:
        st.info("Belum ada data siswa. Upload file Excel di atas.")
        return

    per_class = df.groupby("Kelas").size().reset_index(name="Jumlah siswa")
    f1, f2 = st.columns([1, 2])
    chosen = f1.selectbox("Kelas", [ALL_CLASSES] + per_class["Kelas"].tolist(), key="stu_class")
    query = f2.text_input("Cari nama / NIS", key="stu_query")

    view = df if chosen == ALL_CLASSES else df[df["Kelas"] == chosen]
    if query.strip():
        q = query.strip().lower()
        view = view[view["Nama"].str.lower().str.contains(q, regex=False)
                    | view["NIS"].str.lower().str.contains(q, regex=False)]
    st.caption(f"Menampilkan {len(view)} dari {len(df)} siswa.")
    st.dataframe(view.reset_index(drop=True), hide_index=True)
    with st.expander("Jumlah siswa per kelas"):
        st.dataframe(per_class, hide_index=True)


# ---- Halaman 3: Pemindaian LJK --------------------------------------------------------
def page_scan():
    st.title("📷 Pemindaian LJK")
    quizzes = list_quizzes()
    if not quizzes:
        st.info("Belum ada kuis. Buat kuis dulu di halaman Dashboard.")
        return
    students = list_students()
    if students.empty:
        st.warning("Belum ada data siswa. Hasil scan baru bisa disimpan setelah siswa diimpor.")

    ids = [q["id"] for q in quizzes]
    labels = {q["id"]: f"#{q['id']} · {q['name']} ({q['total_questions']} soal)" for q in quizzes}
    show_flash()
    qid = st.selectbox("Kuis yang diperiksa", ids, key="scan_quiz", format_func=lambda i: labels[i])
    quiz = get_quiz(qid)

    with st.expander("Pengaturan lanjutan"):
        threshold = st.slider("Ambang arsiran (makin kecil makin peka terhadap arsiran tipis)",
                              0.20, 0.70, DEFAULT_FILL_THRESHOLD, 0.05, key="fill_thr")
        show_key = st.checkbox("Tandai juga kunci yang benar pada soal yang salah", key="show_key")

    source = st.radio("Sumber gambar", ["📁 Upload foto", "📷 Kamera"], horizontal=True, key="scan_src")
    ver = st.session_state.setdefault("scan_ver", 0)
    if source.startswith("📁"):
        photo = st.file_uploader("Foto LJK (JPG / PNG)", type=["jpg", "jpeg", "png"], key=f"scan_upl_{ver}")
        st.caption("Di HP, tombol upload menawarkan 'Ambil foto' dengan kamera asli, hasilnya lebih tajam "
                   "daripada mode Kamera di browser.")
    else:
        photo = st.camera_input("Arahkan kamera ke LJK, pastikan keempat kotak hitam sudut terlihat",
                                key=f"scan_cam_{ver}")

    if photo is None:
        st.session_state.pop("scan_sig", None)
        st.session_state.pop("scan", None)
        st.info("Ambil atau upload foto LJK untuk mulai memeriksa.")
        return

    data = photo.getvalue()
    sig = (qid, hashlib.md5(data).hexdigest(), float(threshold), bool(show_key))
    if st.session_state.get("scan_sig") != sig:
        with st.spinner("Memeriksa LJK…"):
            try:
                st.session_state["scan"] = run_scan(data, quiz, threshold, show_key)
                st.session_state["scan_error"] = None
            except (omr.OMRError, ValueError) as exc:
                st.session_state["scan"] = None
                st.session_state["scan_error"] = str(exc)
        st.session_state["scan_sig"] = sig

    scan = st.session_state.get("scan")
    if scan is None:
        st.error(st.session_state.get("scan_error") or "LJK tidak bisa diperiksa.")
        st.caption("Tips: foto seluruh kertas dari atas, cahaya merata tanpa bayangan tajam, keempat kotak "
                   "hitam di sudut harus terlihat utuh.")
        return

    for note in scan["notes"]:
        st.warning(note)
    if scan["mismatch"]:
        st.error(scan["mismatch"])

    # ---- pilih siswa ----
    grade = scan["grade"]
    already = scanned_nis(qid)
    p1, p2 = st.columns(2)
    classes = [ALL_CLASSES] + sorted(students["Kelas"].unique().tolist(), key=str.lower)
    chosen_class = p1.selectbox("Kelas", classes, key="scan_class")
    pool = students if chosen_class == ALL_CLASSES else students[students["Kelas"] == chosen_class]
    names = {r.NIS: f"{r.NIS} — {r.Nama}" + (" ✓" if r.NIS in already else "") for r in pool.itertuples()}
    nis = p2.selectbox("Nama siswa (✓ = sudah punya hasil untuk kuis ini)", list(names), key="scan_student",
                       index=None, placeholder="Pilih siswa…", format_func=lambda v: names[v])
    student_name = None
    if nis is not None:
        student_name = students.loc[students["NIS"] == nis, "Nama"].iloc[0]
        if nis in already:
            st.warning("Siswa ini sudah punya hasil untuk kuis ini. Menyimpan scan baru akan menjadikannya "
                       "hasil terbaru (riwayat lama tetap tersimpan).")

    # ---- hasil koreksi ----
    st.divider()
    m = st.columns(5)
    m[0].metric("Nama siswa", student_name or "— pilih siswa —")
    m[1].metric("Skor total", f"{grade['total_score']:g} / {grade['max_score']:g}")
    m[2].metric("Persentase", f"{grade['persentase']:g}%")
    m[3].metric("Benar / Salah", f"{grade['benar']} / {grade['salah']}")
    m[4].metric("Kosong / Ganda", f"{grade['kosong']} / {grade['ganda']}")
    if grade["ganda"] or grade["kosong"]:
        st.caption("Soal kosong atau diarsir ganda dihitung salah. Jika tampak janggal, periksa pratinjau "
                   "dan foto ulang bila perlu.")

    img_col, table_col = st.columns(2)
    with img_col:
        st.markdown("**Pratinjau LJK** (hijau = benar, merah = salah)")
        st.image(cv2.cvtColor(scan["feedback_bgr"], cv2.COLOR_BGR2RGB))
        ok, buf = cv2.imencode(".png", scan["feedback_bgr"])
        if ok:
            st.download_button("Unduh gambar koreksi", data=buf.tobytes(), file_name="koreksi.png",
                               mime="image/png", key="dl_feedback")
    with table_col:
        st.markdown("**Detail jawaban**")
        only_wrong = st.checkbox("Hanya tampilkan soal yang tidak benar", key="only_wrong")
        table = details_df(grade)
        if only_wrong:
            table = table[~table["Status"].str.startswith("✅")]
        st.dataframe(table, hide_index=True, height=620)

    # ---- simpan ----
    confirm = True
    if scan["mismatch"]:
        confirm = st.checkbox("Saya yakin LJK ini memang untuk kuis yang dipilih; tetap simpan.",
                              key="confirm_mismatch")
    can_save = nis is not None and confirm
    if nis is None:
        st.info("Pilih siswa terlebih dahulu untuk bisa menyimpan hasil.")
    if st.button("💾 Simpan Hasil Scan", type="primary", disabled=not can_save, key="save_scan"):
        try:
            save_scan(scan, quiz, nis, data, getattr(photo, "name", "foto.jpg"))
        except Exception as exc:
            st.error(f"Gagal menyimpan: {exc}")
        else:
            flash("success", f"Hasil {student_name} tersimpan: skor {grade['total_score']:g} "
                             f"({grade['persentase']:g}%). Siap scan LJK berikutnya.")
            st.session_state["scan_ver"] = ver + 1
            for stale in ("scan_sig", "scan", "scan_error"):
                st.session_state.pop(stale, None)
            st.rerun()


# ---- Halaman 4: Laporan & Ekspor ---------------------------------------------------------
def page_report():
    st.title("📈 Laporan & Ekspor Nilai")
    quizzes = list_quizzes()
    if not quizzes:
        st.info("Belum ada kuis.")
        return
    ids = [q["id"] for q in quizzes]
    labels = {q["id"]: f"#{q['id']} · {q['name']} ({q['peserta']} siswa ter-scan)" for q in quizzes}
    r1, r2 = st.columns([3, 1])
    qid = r1.selectbox("Kuis", ids, key="rep_quiz", format_func=lambda i: labels[i])
    kkm = r2.number_input("KKM (%)", min_value=0, max_value=100, value=75, step=1, key="kkm")

    df = results_df(qid)
    if df.empty:
        st.info("Belum ada hasil scan untuk kuis ini.")
        return

    s = score_stats(df, kkm)
    m = st.columns(6)
    m[0].metric("Peserta", s["peserta"])
    m[1].metric("Rata-rata", f"{s['rata']:g}%")
    m[2].metric("Median", f"{s['median']:g}%")
    m[3].metric("Tertinggi", f"{s['maks']:g}%")
    m[4].metric("Terendah", f"{s['min']:g}%")
    m[5].metric(f"Tuntas (≥{kkm}%)", f"{s['lulus']} / {s['peserta']}")
    st.caption(f"Simpangan baku: {s['std']:g}. Yang dihitung adalah hasil scan terbaru tiap siswa.")

    g1, g2 = st.columns(2)
    with g1:
        st.markdown("**Sebaran nilai (%)**")
        st.bar_chart(score_distribution(df))
    with g2:
        st.markdown("**Ringkasan per kelas**")
        st.dataframe(class_summary(df), hide_index=True)

    st.subheader("Daftar nilai")
    st.dataframe(df, hide_index=True)

    quiz = get_quiz(qid)
    base = f"Nilai_{_safe_name(quiz['name'], 'kuis')}"
    x_bytes, _ = export_bytes(qid, "xlsx")
    c_bytes, _ = export_bytes(qid, "csv")
    e1, e2, _ = st.columns([1, 1, 2])
    e1.download_button("📊 Ekspor Nilai ke Excel", data=x_bytes, file_name=base + ".xlsx", type="primary",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_xlsx")
    e2.download_button("Ekspor CSV", data=c_bytes, file_name=base + ".csv", mime="text/csv", key="dl_csv")

    st.subheader("Analisis butir soal")
    items = item_analysis(qid)
    st.bar_chart(items.set_index("No")["% Benar"])
    hardest = items.sort_values(["% Benar", "No"]).head(5)
    st.caption("Soal paling sulit (persentase benar terendah): "
               + ", ".join(f"no. {no} ({pct:g}%)" for no, pct in zip(hardest["No"], hardest["% Benar"])))
    with st.expander("Tabel analisis per soal"):
        st.dataframe(items, hide_index=True)

    everyone = list_students()
    pending = everyone[~everyone["NIS"].isin(df["NIS"])]
    with st.expander(f"Siswa yang belum ter-scan ({len(pending)})"):
        if pending.empty:
            st.write("Semua siswa sudah punya hasil.")
        else:
            st.dataframe(pending.reset_index(drop=True), hide_index=True)


# ---- Navigasi --------------------------------------------------------------------------------
PAGES = {
    "📊 Dashboard & Cetak LJK": page_dashboard,
    "👥 Kelola Data Siswa": page_students,
    "📷 Pemindaian LJK": page_scan,
    "📈 Laporan & Ekspor Nilai": page_report,
}


def main():
    st.set_page_config(page_title="Aplikasi LJK", page_icon="📝", layout="wide")
    dm.init_db()
    st.sidebar.title("📝 Aplikasi LJK")
    page = st.sidebar.radio("Menu", list(PAGES), key="nav")
    st.sidebar.caption(f"Database: {os.path.basename(dm.DB_PATH)}")
    PAGES[page]()


if __name__ == "__main__":
    main()
