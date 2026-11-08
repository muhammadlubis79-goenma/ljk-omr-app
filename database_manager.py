#!/usr/bin/env python3
"""Manajemen database SQLite lokal + integrasi Excel/CSV untuk aplikasi LJK.

Tabel:
    classes       (id, class_name)
    students      (id, student_id, full_name, class_id)
    quizzes       (id, quiz_name, total_questions, answer_key_json, weights_json)
    scan_results  (id, quiz_id, student_id, total_score, student_answers_json,
                   image_path, scanned_at)

Fungsi utama:
    init_db()
    import_students_from_excel(file_path)
    export_grades_to_excel(quiz_id, output_path)
    save_scan_result(quiz_id, student_id, total_score, answers, image_path)

Fungsi pendukung: create_quiz(), calculate_score(), set_db_path().

Dependensi: pip install openpyxl   (sqlite3, csv, json sudah bawaan Python)

Pemakaian cepat lewat terminal:
    python database_manager.py init
    python database_manager.py import-siswa daftar_siswa.xlsx
    python database_manager.py export 1 nilai_kuis_1.xlsx
"""

import argparse
import csv
import io
import json
import os
import re
import sqlite3
from contextlib import contextmanager
from datetime import datetime

# --------------------------------------------------------------------------
# Konfigurasi
# --------------------------------------------------------------------------
# Lokasi file database. Default: ljk_data.db di folder yang sama dengan file ini
# (jadi sama di mana pun script dijalankan). Bisa diganti lewat environment
# variable LJK_DB_PATH atau fungsi set_db_path().
DB_PATH = os.environ.get("LJK_DB_PATH") or os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "ljk_data.db"
)

VALID_OPTIONS = "ABCDE"  # pilihan jawaban yang sah untuk kunci jawaban

# Nama kolom file siswa (dibandingkan tanpa huruf besar/kecil, spasi, underscore).
_HEADER_ALIASES = {
    "nis": {"nis", "nisn", "studentid"},
    "nama": {"namalengkap", "nama", "fullname"},
    "kelas": {"kelas", "class", "classname"},
}
_HEADER_LABELS = {"nis": "NIS", "nama": "Nama_Lengkap", "kelas": "Kelas"}

_SCHEMA = """
CREATE TABLE IF NOT EXISTS classes (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    class_name  TEXT NOT NULL UNIQUE COLLATE NOCASE
);

CREATE TABLE IF NOT EXISTS students (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    student_id  TEXT NOT NULL UNIQUE,
    full_name   TEXT NOT NULL,
    class_id    INTEGER NOT NULL REFERENCES classes(id)
);

CREATE TABLE IF NOT EXISTS quizzes (
    id               INTEGER PRIMARY KEY AUTOINCREMENT,
    quiz_name        TEXT NOT NULL,
    total_questions  INTEGER NOT NULL CHECK (total_questions > 0),
    answer_key_json  TEXT NOT NULL,
    weights_json     TEXT
);

CREATE TABLE IF NOT EXISTS scan_results (
    id                    INTEGER PRIMARY KEY AUTOINCREMENT,
    quiz_id               INTEGER NOT NULL REFERENCES quizzes(id),
    student_id            INTEGER NOT NULL REFERENCES students(id),
    total_score           REAL NOT NULL,
    student_answers_json  TEXT NOT NULL,
    image_path            TEXT,
    scanned_at            TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_students_class ON students(class_id);
CREATE INDEX IF NOT EXISTS idx_scan_quiz_student ON scan_results(quiz_id, student_id);
"""


def set_db_path(path):
    """Ganti lokasi file database (mis. untuk pengujian)."""
    global DB_PATH
    DB_PATH = os.path.abspath(path)


# --------------------------------------------------------------------------
# Koneksi & inisialisasi
# --------------------------------------------------------------------------
@contextmanager
def get_connection():
    """Koneksi SQLite: commit bila sukses, rollback bila error, selalu ditutup."""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    try:
        yield conn
        conn.commit()
    except BaseException:
        conn.rollback()
        raise
    finally:
        conn.close()


def init_db():
    """Buat semua tabel bila belum ada. Aman dipanggil berulang kali."""
    parent = os.path.dirname(DB_PATH)
    if parent:
        os.makedirs(parent, exist_ok=True)
    with get_connection() as conn:
        conn.executescript(_SCHEMA)
    return DB_PATH


# --------------------------------------------------------------------------
# Fungsi bantu: pembersihan & normalisasi data
# --------------------------------------------------------------------------
def _clean(value):
    """Ubah nilai sel apa pun jadi teks rapi (angka bulat Excel 123.0 -> '123')."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return " ".join(str(value).split())


def _normalize_header(name):
    return re.sub(r"[^a-z0-9]", "", _clean(name).lower())


def _normalize_answer_map(data, total_questions, label, is_key=False):
    """Seragamkan jawaban (list atau dict) menjadi {"1": "A", "2": None, ...}.

    - list : indeks 0 = soal nomor 1.
    - dict : kunci = nomor soal (int atau teks angka).
    - Jawaban kosong / None disimpan sebagai None.
    - Kunci jawaban (is_key=True) wajib lengkap dan berisi satu huruf A-E.
    """
    if isinstance(data, (list, tuple)):
        pairs = list(enumerate(data, start=1))
    elif isinstance(data, dict):
        pairs = []
        for k, v in data.items():
            try:
                pairs.append((int(str(k).strip()), v))
            except ValueError:
                raise ValueError(f"Nomor soal pada {label} tidak valid: {k!r}") from None
    else:
        raise ValueError(f"{label} harus berupa list atau dict, bukan {type(data).__name__}")

    result = {str(i): None for i in range(1, total_questions + 1)}
    for number, value in pairs:
        if not 1 <= number <= total_questions:
            raise ValueError(
                f"Nomor soal {number} pada {label} di luar rentang 1-{total_questions}"
            )
        text = _clean(value).upper().replace(" ", "")
        result[str(number)] = text or None

    if is_key:
        for number, value in result.items():
            if value is None or len(value) != 1 or value not in VALID_OPTIONS:
                raise ValueError(
                    f"Kunci jawaban soal {number} harus salah satu dari "
                    f"{', '.join(VALID_OPTIONS)} (ditemukan: {value!r})"
                )
    return result


def _normalize_weights(weights, total_questions):
    """Seragamkan bobot per soal menjadi {"1": 1.0, ...}. None = semua bobot 1."""
    result = {str(i): 1.0 for i in range(1, total_questions + 1)}
    if weights is None:
        return result

    if isinstance(weights, (list, tuple)):
        if len(weights) != total_questions:
            raise ValueError(
                f"Jumlah bobot ({len(weights)}) harus sama dengan jumlah soal ({total_questions})"
            )
        pairs = list(enumerate(weights, start=1))
    elif isinstance(weights, dict):
        pairs = []
        for k, v in weights.items():
            try:
                pairs.append((int(str(k).strip()), v))
            except ValueError:
                raise ValueError(f"Nomor soal pada bobot tidak valid: {k!r}") from None
    else:
        raise ValueError("weights harus berupa list, dict, atau None")

    for number, value in pairs:
        if not 1 <= number <= total_questions:
            raise ValueError(f"Nomor soal {number} pada bobot di luar rentang 1-{total_questions}")
        if isinstance(value, bool) or not isinstance(value, (int, float)) or value < 0 or value != value:
            raise ValueError(f"Bobot soal {number} harus berupa angka >= 0 (ditemukan: {value!r})")
        result[str(number)] = float(value)
    return result


def _load_map(json_text):
    """JSON tersimpan -> dict dengan kunci int: {1: 'A', 2: None, ...}."""
    return {int(k): v for k, v in json.loads(json_text).items()}


def _grade(answer_key, weights, answers):
    """Bandingkan jawaban dengan kunci. Mengembalikan (skor, jumlah_benar).

    Jawaban kosong atau bertanda ganda (mis. 'AB') tidak pernah dianggap benar.
    """
    score = 0.0
    correct = 0
    for number, key in answer_key.items():
        if answers.get(number) == key:
            score += weights.get(number, 1.0)
            correct += 1
    return score, correct


def _get_quiz(conn, quiz_id):
    row = conn.execute("SELECT * FROM quizzes WHERE id = ?", (quiz_id,)).fetchone()
    if row is None:
        raise ValueError(f"Kuis dengan id {quiz_id!r} tidak ditemukan")
    return row


def _get_or_create_class(conn, class_name, new_classes=None):
    """Cari kelas (tanpa membedakan huruf besar/kecil); buat baru bila belum ada."""
    row = conn.execute(
        "SELECT id FROM classes WHERE class_name = ? COLLATE NOCASE", (class_name,)
    ).fetchone()
    if row:
        return row["id"]
    cur = conn.execute("INSERT INTO classes (class_name) VALUES (?)", (class_name,))
    if new_classes is not None:
        new_classes.append(class_name)
    return cur.lastrowid


# --------------------------------------------------------------------------
# Membaca file tabel (.xlsx / .csv)
# --------------------------------------------------------------------------
def _read_xlsx(path):
    from openpyxl import load_workbook

    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb.active
        return [list(row) for row in ws.iter_rows(values_only=True)]
    finally:
        wb.close()


def _read_csv(path):
    with open(path, "rb") as fh:
        raw = fh.read()
    text = None
    for encoding in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            text = raw.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    # Excel versi Indonesia sering menyimpan CSV dengan pemisah titik koma.
    first_line = next((ln for ln in text.splitlines() if ln.strip()), "")
    delimiter = max(",;\t", key=first_line.count) if first_line else ","
    return list(csv.reader(io.StringIO(text), delimiter=delimiter))


def _read_table(file_path):
    if not os.path.isfile(file_path):
        raise FileNotFoundError(f"File tidak ditemukan: {file_path}")
    ext = os.path.splitext(file_path)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        return _read_xlsx(file_path)
    if ext == ".csv":
        return _read_csv(file_path)
    raise ValueError(
        f"Format file '{ext}' tidak didukung. Gunakan .xlsx atau .csv "
        "(file .xls lama: simpan ulang dulu sebagai .xlsx)."
    )


# --------------------------------------------------------------------------
# Fungsi utama 1: impor siswa
# --------------------------------------------------------------------------
def import_students_from_excel(file_path):
    """Impor siswa dari .xlsx / .csv dengan kolom [NIS, Nama_Lengkap, Kelas].

    - Kelas yang belum ada di tabel `classes` dibuat otomatis.
    - NIS yang sudah ada diperbarui (nama dan kelas), bukan digandakan.
    - Baris dengan data tidak lengkap dilewati dan dicatat di `errors`.
    - Seluruh impor satu transaksi: bila terjadi error, tidak ada yang tersimpan.

    Tips: format kolom NIS sebagai Text di Excel agar angka 0 di depan tidak hilang.

    Mengembalikan dict:
        {"inserted": int, "updated": int, "skipped": int,
         "new_classes": [nama_kelas, ...], "errors": [pesan, ...]}
    """
    init_db()
    table = _read_table(file_path)

    header_index = next(
        (i for i, row in enumerate(table) if any(_clean(c) for c in row)), None
    )
    if header_index is None:
        raise ValueError("File kosong: tidak ada header kolom.")

    headers = [_normalize_header(c) for c in table[header_index]]
    columns = {}
    for key, aliases in _HEADER_ALIASES.items():
        for idx, name in enumerate(headers):
            if name in aliases:
                columns[key] = idx
                break
    missing = [_HEADER_LABELS[k] for k in _HEADER_ALIASES if k not in columns]
    if missing:
        found = ", ".join(_clean(c) for c in table[header_index] if _clean(c)) or "(kosong)"
        raise ValueError(
            f"Kolom wajib tidak ditemukan: {', '.join(missing)}. Kolom yang terbaca: {found}"
        )

    result = {"inserted": 0, "updated": 0, "skipped": 0, "new_classes": [], "errors": []}

    with get_connection() as conn:
        for row_number, row in enumerate(table[header_index + 1:], start=header_index + 2):
            def cell(key):
                idx = columns[key]
                return _clean(row[idx]) if idx < len(row) else ""

            nis, nama, kelas = cell("nis"), cell("nama"), cell("kelas")

            if not (nis or nama or kelas):
                continue  # baris kosong, abaikan diam-diam
            if not (nis and nama and kelas):
                result["skipped"] += 1
                result["errors"].append(
                    f"Baris {row_number}: data tidak lengkap "
                    f"(NIS={nis!r}, Nama={nama!r}, Kelas={kelas!r})"
                )
                continue

            class_id = _get_or_create_class(conn, kelas, result["new_classes"])
            existing = conn.execute(
                "SELECT id FROM students WHERE student_id = ?", (nis,)
            ).fetchone()
            if existing:
                conn.execute(
                    "UPDATE students SET full_name = ?, class_id = ? WHERE id = ?",
                    (nama, class_id, existing["id"]),
                )
                result["updated"] += 1
            else:
                conn.execute(
                    "INSERT INTO students (student_id, full_name, class_id) VALUES (?, ?, ?)",
                    (nis, nama, class_id),
                )
                result["inserted"] += 1

    return result


# --------------------------------------------------------------------------
# Kuis & penilaian (pendukung)
# --------------------------------------------------------------------------
def create_quiz(quiz_name, total_questions, answer_key, weights=None):
    """Buat kuis baru dan kembalikan id-nya.

    answer_key : list ["A","C",...] (indeks 0 = soal 1) atau dict {1: "A", ...};
                 wajib lengkap untuk semua soal, nilainya A-E.
    weights    : None (semua soal bobot 1), list, atau dict bobot per soal.
    """
    init_db()
    quiz_name = _clean(quiz_name)
    if not quiz_name:
        raise ValueError("quiz_name tidak boleh kosong")
    if (not isinstance(total_questions, int) or isinstance(total_questions, bool)
            or total_questions < 1):
        raise ValueError("total_questions harus bilangan bulat >= 1")

    key_map = _normalize_answer_map(answer_key, total_questions, "kunci jawaban", is_key=True)
    weight_map = _normalize_weights(weights, total_questions)

    with get_connection() as conn:
        cur = conn.execute(
            "INSERT INTO quizzes (quiz_name, total_questions, answer_key_json, weights_json) "
            "VALUES (?, ?, ?, ?)",
            (quiz_name, total_questions,
             json.dumps(key_map, ensure_ascii=False), json.dumps(weight_map)),
        )
        return cur.lastrowid


def calculate_score(quiz_id, answers):
    """Hitung nilai dari jawaban siswa terhadap kunci kuis (memperhitungkan bobot).

    Hasilnya bisa dipakai sebagai `total_score` untuk save_scan_result().
    Mengembalikan dict: total_score, max_score, benar, salah, kosong, persentase.
    `salah` = soal yang tidak benar (termasuk kosong dan tanda ganda).
    """
    init_db()
    with get_connection() as conn:
        quiz = _get_quiz(conn, quiz_id)
    total = quiz["total_questions"]
    key = _load_map(quiz["answer_key_json"])
    weights = _load_map(quiz["weights_json"])
    answer_map = _load_map(
        json.dumps(_normalize_answer_map(answers, total, "jawaban siswa"))
    )

    score, correct = _grade(key, weights, answer_map)
    max_score = sum(weights.values())
    blank = sum(1 for v in answer_map.values() if v is None)
    return {
        "total_score": round(score, 2),
        "max_score": round(max_score, 2),
        "benar": correct,
        "salah": total - correct,
        "kosong": blank,
        "persentase": round(score / max_score * 100, 2) if max_score > 0 else 0.0,
    }


# --------------------------------------------------------------------------
# Fungsi utama 2: simpan hasil scan
# --------------------------------------------------------------------------
def save_scan_result(quiz_id, student_id, total_score, answers, image_path=None):
    """Simpan hasil pemeriksaan satu LJK. Mengembalikan id baris scan_results.

    quiz_id     : id kuis (kolom quizzes.id).
    student_id  : NIS siswa (kolom students.student_id), BUKAN id internal.
    total_score : skor akhir; harus antara 0 dan skor maksimum kuis.
    answers     : list ["A", None, "C", ...] (indeks 0 = soal 1) atau dict {1: "A", ...}.
    image_path  : lokasi foto LJK hasil scan (opsional).

    Scan ulang untuk siswa yang sama pada kuis yang sama disimpan sebagai baris
    baru (riwayat tetap ada); ekspor nilai memakai scan terbaru.
    """
    init_db()
    with get_connection() as conn:
        quiz = _get_quiz(conn, quiz_id)

        nis = _clean(student_id)
        student = conn.execute(
            "SELECT id FROM students WHERE student_id = ?", (nis,)
        ).fetchone()
        if student is None:
            raise ValueError(f"Siswa dengan NIS {nis!r} tidak ditemukan. Impor data siswa dulu.")

        answer_map = _normalize_answer_map(
            answers, quiz["total_questions"], "jawaban siswa"
        )

        max_score = sum(_load_map(quiz["weights_json"]).values())
        try:
            score = float(total_score)
        except (TypeError, ValueError):
            raise ValueError(f"total_score harus berupa angka (ditemukan: {total_score!r})") from None
        if score != score or score < 0 or score > max_score + 1e-9:
            raise ValueError(
                f"total_score {total_score!r} di luar rentang 0 sampai {max_score:g}"
            )

        cur = conn.execute(
            "INSERT INTO scan_results "
            "(quiz_id, student_id, total_score, student_answers_json, image_path, scanned_at) "
            "VALUES (?, ?, ?, ?, ?, ?)",
            (
                quiz["id"],
                student["id"],
                round(score, 2),
                json.dumps(answer_map, ensure_ascii=False),
                str(image_path) if image_path else None,
                datetime.now().isoformat(timespec="seconds"),
            ),
        )
        return cur.lastrowid


# --------------------------------------------------------------------------
# Fungsi utama 3: ekspor nilai
# --------------------------------------------------------------------------
_EXPORT_HEADER = ["NIS", "Nama", "Kelas", "Total Skor", "Benar", "Salah", "Persentase"]


def _write_xlsx(path, records):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "Nilai"
    ws.append(_EXPORT_HEADER)
    for record in records:
        ws.append(record)

    header_fill = PatternFill("solid", start_color="305496", end_color="305496")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row in ws.iter_rows(min_row=2):
        row[0].number_format = "@"                 # NIS tetap teks
        row[6].number_format = '0.00"%"'           # 86.67 tampil sebagai 86.67%

    for idx, header in enumerate(_EXPORT_HEADER, start=1):
        longest = max(
            [len(str(header))] + [len(str(r[idx - 1])) for r in records]
        )
        ws.column_dimensions[get_column_letter(idx)].width = min(max(longest + 2, 10), 40)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(path)


def _write_csv(path, records):
    with open(path, "w", newline="", encoding="utf-8-sig") as fh:
        writer = csv.writer(fh)
        writer.writerow(_EXPORT_HEADER)
        writer.writerows(records)


def export_grades_to_excel(quiz_id, output_path):
    """Ekspor hasil scan satu kuis ke file .xlsx (atau .csv).

    Kolom: NIS, Nama, Kelas, Total Skor, Benar, Salah, Persentase.
      - Benar      : jumlah jawaban yang cocok dengan kunci.
      - Salah      : jumlah soal - Benar (soal kosong dan tanda ganda ikut dihitung salah).
      - Persentase : Total Skor / skor maksimum kuis x 100 (dibulatkan 2 desimal).
    Bila satu siswa discan lebih dari sekali, hanya scan terbaru yang diekspor.
    Urutan: Kelas, lalu Nama. Mengembalikan jumlah baris data yang ditulis.
    """
    init_db()
    ext = os.path.splitext(output_path)[1].lower()
    if ext not in (".xlsx", ".csv"):
        raise ValueError("output_path harus berakhiran .xlsx atau .csv")

    with get_connection() as conn:
        quiz = _get_quiz(conn, quiz_id)
        total = quiz["total_questions"]
        key = _load_map(quiz["answer_key_json"])
        weights = _load_map(quiz["weights_json"])
        max_score = sum(weights.values())

        rows = conn.execute(
            """
            SELECT s.student_id AS nis, s.full_name AS nama, c.class_name AS kelas,
                   r.total_score AS total_score,
                   r.student_answers_json AS answers_json
            FROM scan_results r
            JOIN students s ON s.id = r.student_id
            JOIN classes  c ON c.id = s.class_id
            WHERE r.quiz_id = ?
              AND r.id = (SELECT MAX(r2.id) FROM scan_results r2
                          WHERE r2.quiz_id = r.quiz_id AND r2.student_id = r.student_id)
            ORDER BY c.class_name COLLATE NOCASE, s.full_name COLLATE NOCASE
            """,
            (quiz["id"],),
        ).fetchall()

    records = []
    for row in rows:
        _, correct = _grade(key, weights, _load_map(row["answers_json"]))
        percent = round(row["total_score"] / max_score * 100, 2) if max_score > 0 else 0.0
        records.append([
            row["nis"], row["nama"], row["kelas"],
            row["total_score"], correct, total - correct, percent,
        ])

    parent = os.path.dirname(os.path.abspath(output_path))
    os.makedirs(parent, exist_ok=True)
    try:
        if ext == ".xlsx":
            _write_xlsx(output_path, records)
        else:
            _write_csv(output_path, records)
    except PermissionError:
        raise PermissionError(
            f"Tidak bisa menulis ke {output_path}. Jika file itu sedang terbuka di Excel, "
            "tutup dulu lalu ulangi."
        ) from None
    return len(records)


# --------------------------------------------------------------------------
# Command line (opsional)
# --------------------------------------------------------------------------
def _main():
    parser = argparse.ArgumentParser(description="Manajemen database LJK (SQLite)")
    sub = parser.add_subparsers(dest="command", required=True)

    sub.add_parser("init", help="buat tabel database")

    p_import = sub.add_parser("import-siswa", help="impor siswa dari .xlsx / .csv")
    p_import.add_argument("file")

    p_export = sub.add_parser("export", help="ekspor nilai satu kuis")
    p_export.add_argument("quiz_id", type=int)
    p_export.add_argument("output")

    args = parser.parse_args()

    if args.command == "init":
        print("Database siap:", init_db())
    elif args.command == "import-siswa":
        res = import_students_from_excel(args.file)
        print(f"Baru: {res['inserted']}  Diperbarui: {res['updated']}  Dilewati: {res['skipped']}")
        if res["new_classes"]:
            print("Kelas baru:", ", ".join(res["new_classes"]))
        for msg in res["errors"]:
            print("  !", msg)
    elif args.command == "export":
        count = export_grades_to_excel(args.quiz_id, args.output)
        print(f"{count} baris diekspor ke {args.output}")


if __name__ == "__main__":
    _main()
