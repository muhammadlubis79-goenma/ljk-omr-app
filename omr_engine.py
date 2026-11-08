#!/usr/bin/env python3
"""Mesin OMR: membaca foto LJK buatan ljk_generator.py dan menilainya otomatis.

Alur:
    detect_and_warp()          foto -> kertas lurus (A4, 8 piksel per mm)
    get_bubble_coordinates()   koordinat tiap bulatan A-E dari rasio ukuran kertas
    process_bubbles()          ukur kehitaman tiap bulatan -> jawaban siswa
    grade_paper()              jawaban siswa vs kunci -> skor berbobot
    draw_visual_feedback()     lingkaran HIJAU (benar) / MERAH (salah) di gambar
    scan_ljk()                 semua langkah di atas dalam satu panggilan

Pemakaian cepat:
    python omr_engine.py foto_ljk.jpg --key ABCDEABCDEABCDE --out hasil.png

Dari kode:
    from omr_engine import scan_ljk
    hasil = scan_ljk("foto_ljk.jpg", answer_key=["A", "B", "C", ...])
    hasil["grade"]["total_score"], hasil["student_answers"]

Hasilnya cocok langsung dengan database_manager.save_scan_result():
    save_scan_result(quiz_id, nis, hasil["grade"]["total_score"],
                     hasil["student_answers"], image_path)

Dependensi: pip install opencv-python numpy
"""

import argparse
import json
import math
import os
import sys
import warnings
from itertools import combinations

import cv2
import numpy as np

# --------------------------------------------------------------------------
# Geometri LJK (milimeter). HARUS sama dengan konstanta di ljk_generator.py:
# bila Anda mengubah layout di generator, ubah juga bagian ini.
# --------------------------------------------------------------------------
PAGE_W_MM, PAGE_H_MM = 210.0, 297.0
MARKER_SIZE_MM = 10.0        # kotak jangkar 1 x 1 cm
MARKER_MARGIN_MM = 5.0       # jarak kotak jangkar dari tepi kertas
SIDE_MARGIN_MM = 20.0
OPTIONS = "ABCDE"
MAX_SOAL = 100
MAX_ROWS_PER_COL = 25
BUBBLE_R_MM = 2.3            # jari-jari bulatan
GRID_TOP_MM = 73.0           # awal blok bulatan, diukur dari tepi ATAS kertas
GRID_BOTTOM_MM = 25.0        # batas bawah blok bulatan, diukur dari tepi BAWAH
GRID_HEADER_MM = 7.0         # tinggi baris huruf A-E
LABEL_W_MM = 11.0            # lebar kolom nomor soal
MAX_PITCH_Y_MM = 8.5
MAX_PITCH_X_MM = 8.0
MAX_COL_W_MM = 60.0
QR_SIZE_MM = 24.0

# --------------------------------------------------------------------------
# Parameter citra
# --------------------------------------------------------------------------
SCALE = 8                                    # piksel per mm pada kertas hasil luruskan
WARP_W = int(round(PAGE_W_MM * SCALE))       # 1680
WARP_H = int(round(PAGE_H_MM * SCALE))       # 2376
MAX_DETECT_SIDE = 1400                       # sisi terpanjang saat mencari kotak jangkar
SAMPLE_RATIO = 0.60                          # bagian dalam bulatan yang diukur (hindari garis tepi)
ALIGN_MIN_SCORE = 0.10                       # skor kecocokan tata letak minimum (lihat _alignment_score)

# Kotak pencarian QR code pada kertas lurus (mm): x0, y0, x1, y1 dari pojok kiri atas
QR_ROI_MM = (159.0, 12.0, 194.0, 48.0)

GREEN = (0, 190, 0)      # BGR
RED = (0, 0, 230)


class OMRError(Exception):
    """Kesalahan pemrosesan LJK (kertas tidak terdeteksi, gambar rusak, dsb.)."""


# --------------------------------------------------------------------------
# Fungsi bantu umum
# --------------------------------------------------------------------------
def _check_total(total_questions):
    if (not isinstance(total_questions, (int, np.integer)) or isinstance(total_questions, bool)
            or not 1 <= int(total_questions) <= MAX_SOAL):
        raise ValueError(f"total_questions harus bilangan bulat 1 sampai {MAX_SOAL}")
    return int(total_questions)


def _load_image(image_or_path):
    """Terima path file atau array BGR/gray. Aman untuk path berhuruf non-ASCII."""
    if isinstance(image_or_path, np.ndarray):
        img = image_or_path
        if img.ndim == 2:
            img = cv2.cvtColor(img, cv2.COLOR_GRAY2BGR)
        if img.ndim != 3 or img.shape[2] not in (3, 4):
            raise OMRError("Array gambar harus berbentuk (tinggi, lebar) atau (tinggi, lebar, 3)")
        return img[:, :, :3].copy()
    path = os.fspath(image_or_path)
    if not os.path.isfile(path):
        raise FileNotFoundError(f"File gambar tidak ditemukan: {path}")
    img = cv2.imdecode(np.fromfile(path, dtype=np.uint8), cv2.IMREAD_COLOR)
    if img is None:
        raise OMRError(f"File bukan gambar yang bisa dibaca: {path}")
    return img


def _to_bgr(img):
    if img.ndim == 2:
        return cv2.cvtColor(img, cv2.COLOR_GRAY2BGR)
    return img[:, :, :3].copy()


def _canonical(warped_image):
    """Pastikan gambar berukuran WARP_W x WARP_H (resize bila rasionya cocok)."""
    if not isinstance(warped_image, np.ndarray) or warped_image.ndim not in (2, 3):
        raise OMRError("warped_image harus berupa array gambar OpenCV")
    h, w = warped_image.shape[:2]
    if (w, h) == (WARP_W, WARP_H):
        return warped_image
    if abs((w / h) / (PAGE_W_MM / PAGE_H_MM) - 1) > 0.03:
        raise OMRError(
            f"Rasio gambar {w}x{h} tidak cocok dengan A4 portrait. "
            "Gunakan keluaran detect_and_warp()."
        )
    return cv2.resize(warped_image, (WARP_W, WARP_H), interpolation=cv2.INTER_CUBIC)


# --------------------------------------------------------------------------
# 1. Deteksi kotak jangkar & luruskan kertas
# --------------------------------------------------------------------------
def _binarize_variants(gray):
    """Dua versi biner (global Otsu dan adaptif) supaya tahan bayangan."""
    h, w = gray.shape
    blur = cv2.GaussianBlur(gray, (5, 5), 0)
    _, otsu = cv2.threshold(blur, 0, 255, cv2.THRESH_BINARY_INV | cv2.THRESH_OTSU)
    block = max(int(min(h, w) * 0.10) | 1, 31)          # harus ganjil
    adaptive = cv2.adaptiveThreshold(
        blur, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, block, 12
    )
    return [otsu, adaptive]


def _marker_candidates(binary, gray):
    """Cari kontur yang berbentuk persegi hitam pekat."""
    h, w = gray.shape
    img_area = float(h * w)
    p95 = float(np.percentile(gray, 95))
    contours, _ = cv2.findContours(binary, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
    found = []
    for cnt in contours:
        area = cv2.contourArea(cnt)
        if not (2e-4 * img_area <= area <= 6e-3 * img_area):
            continue
        (_, _), (rw, rh), _ = cv2.minAreaRect(cnt)
        if min(rw, rh) < 4:
            continue
        if max(rw, rh) / min(rw, rh) > 1.45:             # harus mendekati persegi
            continue
        if area / (rw * rh) < 0.82:                      # lingkaran = 0.785 -> tertolak
            continue
        hull_area = cv2.contourArea(cv2.convexHull(cnt))
        if hull_area <= 0 or area / hull_area < 0.90:
            continue

        x, y, bw, bh = cv2.boundingRect(cnt)             # wajib hitam pekat di dalamnya
        mask = np.zeros((bh, bw), np.uint8)
        cv2.drawContours(mask, [cnt - [x, y]], -1, 255, -1)
        inside = cv2.mean(gray[y:y + bh, x:x + bw], mask)[0]
        if inside > 0.5 * p95:
            continue

        m = cv2.moments(cnt)
        if m["m00"] == 0:
            continue
        found.append({
            "center": (m["m10"] / m["m00"], m["m01"] / m["m00"]),
            "area": area,
            "side": math.sqrt(area),
        })
    return found


def _order_points(pts):
    """Urutkan 4 titik menjadi [kiri-atas, kanan-atas, kanan-bawah, kiri-bawah]."""
    s = pts.sum(axis=1)
    d = np.diff(pts, axis=1).ravel()                     # y - x
    ordered = np.array(
        [pts[np.argmin(s)], pts[np.argmin(d)], pts[np.argmax(s)], pts[np.argmax(d)]],
        dtype=np.float32,
    )
    if len({tuple(p) for p in ordered.tolist()}) < 4:
        return None
    return ordered


def _best_quad(candidates):
    """Pilih 4 kandidat yang membentuk segi empat konveks terbesar dengan ukuran seragam."""
    top = sorted(candidates, key=lambda c: -c["area"])[:14]
    best, best_area = None, 0.0
    for combo in combinations(top, 4):
        sides = np.array([c["side"] for c in combo])
        if sides.std() / sides.mean() > 0.30:
            continue
        pts = np.array([c["center"] for c in combo], dtype=np.float32)
        ordered = _order_points(pts)
        if ordered is None or not cv2.isContourConvex(ordered.astype(np.int32)):
            continue
        tl, tr, br, bl = ordered
        width = (np.linalg.norm(tr - tl) + np.linalg.norm(br - bl)) / 2
        height = (np.linalg.norm(bl - tl) + np.linalg.norm(br - tr)) / 2
        ratio = width / height
        if not (0.45 <= ratio <= 0.95 or 1.05 <= ratio <= 2.2):   # kertas A4: 0.69 atau 1.46
            continue
        area = cv2.contourArea(ordered)
        if area > best_area:
            best, best_area = ordered, area
    return best


def _find_marker_points(img):
    """Mengembalikan 4 titik pusat kotak jangkar (urutan TL, TR, BR, BL) pada gambar asli."""
    h0, w0 = img.shape[:2]
    scale = min(1.0, MAX_DETECT_SIDE / max(h0, w0))
    small = cv2.resize(img, None, fx=scale, fy=scale, interpolation=cv2.INTER_AREA) if scale < 1 else img
    gray = cv2.cvtColor(small, cv2.COLOR_BGR2GRAY)

    most = 0
    for binary in _binarize_variants(gray):
        cands = _marker_candidates(binary, gray)
        most = max(most, len(cands))
        quad = _best_quad(cands) if len(cands) >= 4 else None
        if quad is not None:
            return quad / scale
    raise OMRError(
        f"Hanya {most} kotak hitam penanda yang terdeteksi (butuh 4). Pastikan keempat sudut "
        "kertas terlihat utuh di foto, pencahayaan merata, dan kertas tidak terpotong."
    )


def _decode_qr(warped):
    """Baca teks QR code dari area QR pada kertas lurus. Kosong bila gagal."""
    x0, y0, x1, y1 = (int(round(v * SCALE)) for v in QR_ROI_MM)
    roi = warped[y0:y1, x0:x1]
    gray = cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY) if roi.ndim == 3 else roi

    # Beberapa versi gambar (asli, dipertajam, diperbesar) agar QR yang agak buram tetap terbaca
    blur = cv2.GaussianBlur(gray, (0, 0), 2.0)
    variants = (
        gray,
        cv2.addWeighted(gray, 2.0, blur, -1.0, 0),
        cv2.resize(gray, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC),
        cv2.resize(gray, None, fx=3, fy=3, interpolation=cv2.INTER_CUBIC),
    )
    detectors = [cv2.QRCodeDetector()]
    if hasattr(cv2, "QRCodeDetectorAruco"):
        detectors.append(cv2.QRCodeDetectorAruco())
    for variant in variants:
        for detector in detectors:
            try:
                text, _, _ = detector.detectAndDecode(variant)
            except cv2.error:
                continue
            if text:
                return text
    return ""


def read_qr_metadata(warped_image):
    """Baca metadata QR ({"Quiz_ID", "Jumlah_Soal", "Tanggal"}) dari kertas lurus, atau None."""
    text = _decode_qr(_to_bgr(_canonical(warped_image)))
    if not text:
        return None
    try:
        data = json.loads(text)
    except ValueError:
        return None
    return data if isinstance(data, dict) else None


def _alignment_score(gray, total_questions):
    """Seberapa cocok pola cincin bulatan pada gambar dengan tata letak LJK (0 - 1).

    Pada tiap bulatan dibandingkan kecerahan pita di luar cincin dengan kecerahan cincinnya,
    dicoba pada beberapa pergeseran kecil (+-6 px) agar kertas yang sedikit melengkung tetap
    dihitung cocok. Kertas lurus & berarah benar memberi nilai tinggi (~0.15-0.5); terbalik,
    salah jumlah soal, atau bukan LJK memberi nilai mendekati 0. Isian pensil tidak
    mengganggu karena hanya cincin dan pita luarnya yang diukur.
    """
    flat = _flatten_illumination(gray).astype(np.float32)
    h, w = flat.shape
    radius = BUBBLE_R_MM * SCALE
    angles = np.linspace(0, 2 * math.pi, 24, endpoint=False)

    def offsets(radii):
        dy = np.concatenate([np.rint(r * np.sin(angles)) for r in radii]).astype(int)
        dx = np.concatenate([np.rint(r * np.cos(angles)) for r in radii]).astype(int)
        return dy, dx

    ring_dy, ring_dx = offsets((radius - 1, radius, radius + 1))
    out_dy, out_dx = offsets((radius + 4, radius + 6))
    centers = np.array([c for cs in get_bubble_coordinates(total_questions).values() for c in cs])
    cx, cy = centers[:, 0:1], centers[:, 1:2]

    best = np.full(len(centers), -np.inf)
    for sy in (-6, -3, 0, 3, 6):
        for sx in (-6, -3, 0, 3, 6):
            ring = flat[np.clip(cy + sy + ring_dy, 0, h - 1), np.clip(cx + sx + ring_dx, 0, w - 1)].mean(axis=1)
            outer = flat[np.clip(cy + sy + out_dy, 0, h - 1), np.clip(cx + sx + out_dx, 0, w - 1)].mean(axis=1)
            best = np.maximum(best, outer - ring)
    return float(best.mean()) / 255.0


def detect_and_warp(image_path, total_questions=None, auto_orient=True, debug_path=None):
    """Deteksi 4 kotak jangkar lalu luruskan kertas menjadi A4 portrait (1680 x 2376 px).

    image_path      : path foto, atau array gambar OpenCV (mis. frame kamera).
    total_questions : (opsional) jumlah soal pada LJK. Dipakai sebagai cadangan penentu arah
                      bila QR code tidak terbaca: arah dipilih dari kecocokan pola bulatan.
    auto_orient     : bila True, foto yang terbalik / diputar 90 derajat diperbaiki otomatis
                      (utama: baca QR di pojok kanan atas; cadangan: pola bulatan).
    debug_path      : (opsional) simpan foto asli dengan penanda terdeteksi, untuk diagnosis.

    Mengembalikan gambar BGR kertas yang sudah lurus. Melempar OMRError bila gagal.
    """
    img = _load_image(image_path)
    points = _find_marker_points(img)

    if debug_path:
        dbg = img.copy()
        thick = max(2, img.shape[1] // 300)
        cv2.polylines(dbg, [points.astype(np.int32)], True, (0, 255, 0), thick)
        for i, p in enumerate(points):
            cv2.circle(dbg, (int(p[0]), int(p[1])), thick * 3, (0, 0, 255), -1)
            cv2.putText(dbg, "TL TR BR BL".split()[i], (int(p[0]) + 10, int(p[1]) - 10),
                        cv2.FONT_HERSHEY_SIMPLEX, max(1.0, img.shape[1] / 1500), (255, 0, 0), thick)
        cv2.imencode(os.path.splitext(debug_path)[1] or ".png", dbg)[1].tofile(debug_path)

    # Posisi pusat kotak jangkar pada kertas lurus
    c = MARKER_MARGIN_MM + MARKER_SIZE_MM / 2
    dst = np.array([[c, c], [PAGE_W_MM - c, c], [PAGE_W_MM - c, PAGE_H_MM - c], [c, PAGE_H_MM - c]],
                   dtype=np.float32) * SCALE

    # Kertas tampak portrait (sisi tegak lebih panjang) atau landscape (foto diputar 90 derajat)
    tl, tr, br, bl = points
    width = (np.linalg.norm(tr - tl) + np.linalg.norm(br - bl)) / 2
    height = (np.linalg.norm(bl - tl) + np.linalg.norm(br - tr)) / 2
    base = 0 if width <= height else 1
    orderings = [np.roll(points, -base, axis=0), np.roll(points, -(base + 2), axis=0)]

    warps = []
    for ordered in orderings:
        matrix = cv2.getPerspectiveTransform(ordered.astype(np.float32), dst)
        warped = cv2.warpPerspective(
            img, matrix, (WARP_W, WARP_H), flags=cv2.INTER_CUBIC,
            borderMode=cv2.BORDER_CONSTANT, borderValue=(255, 255, 255),
        )
        if not auto_orient:
            return warped
        if _decode_qr(warped):                       # QR terbaca = arah pasti benar
            return warped
        warps.append(warped)

    # QR tidak terbaca: pilih arah dengan pola bulatan yang paling cocok
    if total_questions is not None:
        n = _check_total(total_questions)
        scores = [_alignment_score(cv2.cvtColor(w, cv2.COLOR_BGR2GRAY), n) for w in warps]
        best = int(np.argmax(scores))
        if scores[best] >= ALIGN_MIN_SCORE and scores[best] > 1.5 * scores[1 - best]:
            return warps[best]

    warnings.warn(
        "QR code tidak terbaca dan arah kertas tidak bisa dipastikan; mengasumsikan foto tegak. "
        "Jika hasil bacaan kacau, foto ulang dengan kertas tegak dan QR terlihat jelas.",
        RuntimeWarning, stacklevel=2,
    )
    return warps[0]


# --------------------------------------------------------------------------
# 2. Koordinat bulatan & pembacaan arsiran
# --------------------------------------------------------------------------
def get_bubble_coordinates(total_questions):
    """Koordinat piksel pusat bulatan pada kertas lurus.

    Mengembalikan {nomor_soal: [(x, y) untuk A, B, C, D, E]} (nomor soal mulai dari 1).
    Rumus mengikuti tata letak ljk_generator.py (rasio terhadap ukuran A4).
    """
    n = _check_total(total_questions)
    n_cols = -(-n // MAX_ROWS_PER_COL)
    rows = -(-n // n_cols)

    grid_bottom = PAGE_H_MM - GRID_BOTTOM_MM
    pitch_y = min(MAX_PITCH_Y_MM, (grid_bottom - GRID_TOP_MM - GRID_HEADER_MM) / rows)
    col_w = min((PAGE_W_MM - 2 * SIDE_MARGIN_MM) / n_cols, MAX_COL_W_MM)
    pitch_x = min(MAX_PITCH_X_MM, (col_w - LABEL_W_MM - 1.0) / len(OPTIONS))
    x_start = (PAGE_W_MM - n_cols * col_w) / 2

    coords = {}
    for q in range(1, n + 1):
        col, row = divmod(q - 1, rows)               # nomor soal berurut per kolom
        y = GRID_TOP_MM + GRID_HEADER_MM + pitch_y * (row + 0.5)
        coords[q] = [
            (int(round((x_start + col * col_w + LABEL_W_MM + pitch_x * (j + 0.5)) * SCALE)),
             int(round(y * SCALE)))
            for j in range(len(OPTIONS))
        ]
    return coords


def _flatten_illumination(gray):
    """Hilangkan bayangan/gradasi cahaya: bagi gambar dengan estimasi warna kertas."""
    small = cv2.resize(gray, None, fx=0.25, fy=0.25, interpolation=cv2.INTER_AREA)
    kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (21, 21))   # ~84 px, > bulatan terisi
    background = cv2.morphologyEx(small, cv2.MORPH_CLOSE, kernel)
    background = cv2.GaussianBlur(background, (0, 0), 3)
    background = cv2.resize(background, (gray.shape[1], gray.shape[0]), interpolation=cv2.INTER_LINEAR)
    return cv2.divide(gray, np.maximum(background, 1), scale=255)


def process_bubbles(warped_image, total_questions, fill_threshold=0.40,
                    multi_mark_ratio=0.65, binary_threshold=150):
    """Baca bulatan yang diarsir pada kertas lurus.

    warped_image     : keluaran detect_and_warp().
    total_questions  : jumlah soal pada LJK (1-100).
    fill_threshold   : bulatan dianggap diarsir bila persentase piksel hitam di bagian
                       dalamnya >= nilai ini (0-1).
    multi_mark_ratio : bulatan lain ikut dianggap diarsir bila kehitamannya >= rasio ini
                       terhadap bulatan tergelap di soal yang sama. Menangani jawaban yang
                       dihapus (bekas samar) tetapi menandai arsiran ganda yang sama pekat.
    binary_threshold : ambang biner (0-255) setelah cahaya diratakan; None = Otsu.

    Mengembalikan (student_answers, fill_ratios):
      student_answers : list panjang total_questions; tiap isi "A".."E", None (kosong),
                        atau gabungan huruf seperti "AC" (arsiran ganda, tidak pernah benar).
      fill_ratios     : list of list persentase kehitaman 0-1 untuk A..E tiap soal.
    """
    n = _check_total(total_questions)
    image = _canonical(warped_image)
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY) if image.ndim == 3 else image

    flat = cv2.GaussianBlur(_flatten_illumination(gray), (3, 3), 0)
    if binary_threshold is None:
        otsu, _ = cv2.threshold(flat, 0, 255, cv2.THRESH_BINARY_INV | cv2.THRESH_OTSU)
        binary_threshold = float(np.clip(otsu, 120, 190))
    _, binary = cv2.threshold(flat, binary_threshold, 255, cv2.THRESH_BINARY_INV)

    r = max(2, int(round(BUBBLE_R_MM * SCALE * SAMPLE_RATIO)))
    disk = np.zeros((2 * r + 1, 2 * r + 1), np.uint8)
    cv2.circle(disk, (r, r), r, 1, -1)
    disk_px = float(disk.sum())

    answers, all_ratios = [], []
    for q, centers in get_bubble_coordinates(n).items():
        ratios = []
        for (cx, cy) in centers:
            patch = binary[cy - r:cy + r + 1, cx - r:cx + r + 1]
            if patch.shape != disk.shape:
                ratios.append(0.0)
                continue
            ratios.append(float((patch > 0)[disk > 0].sum() / disk_px))
        all_ratios.append(ratios)

        top = max(ratios)
        if top < fill_threshold:
            answers.append(None)
            continue
        marked = [OPTIONS[j] for j, v in enumerate(ratios)
                  if v >= fill_threshold and v >= multi_mark_ratio * top]
        answers.append("".join(marked))
    return answers, all_ratios


# --------------------------------------------------------------------------
# 3. Penilaian
# --------------------------------------------------------------------------
def _normalize_answers(data, total, label, is_key=False):
    """list / dict / string -> list panjang `total` berisi huruf kapital atau None."""
    if isinstance(data, np.ndarray):
        data = data.tolist()
    if isinstance(data, str):
        data = list(data)
    if isinstance(data, (list, tuple)):
        if len(data) > total:
            raise ValueError(f"{label} berisi {len(data)} soal, padahal jumlah soal {total}")
        pairs = list(enumerate(data, start=1))
    elif isinstance(data, dict):
        pairs = []
        for k, v in data.items():
            try:
                pairs.append((int(str(k).strip()), v))
            except ValueError:
                raise ValueError(f"Nomor soal pada {label} tidak valid: {k!r}") from None
    else:
        raise ValueError(f"{label} harus berupa list, dict, atau string")

    result = [None] * total
    for number, value in pairs:
        if not 1 <= number <= total:
            raise ValueError(f"Nomor soal {number} pada {label} di luar rentang 1-{total}")
        text = "" if value is None else str(value).strip().upper().replace(" ", "")
        result[number - 1] = text or None

    if is_key:
        for i, value in enumerate(result, start=1):
            if value is None or len(value) != 1 or value not in OPTIONS:
                raise ValueError(
                    f"Kunci jawaban soal {i} harus salah satu dari {', '.join(OPTIONS)} "
                    f"(ditemukan: {value!r})"
                )
    return result


def _infer_total(answer_key):
    if isinstance(answer_key, np.ndarray):
        answer_key = answer_key.tolist()
    if isinstance(answer_key, (list, tuple, str)):
        return len(answer_key)
    if isinstance(answer_key, dict) and answer_key:
        try:
            return max(int(str(k).strip()) for k in answer_key)
        except ValueError:
            pass
    raise ValueError("answer_key harus berupa list, dict, atau string yang tidak kosong")


def _normalize_weights(weights, total):
    if weights is None:
        return [1.0] * total
    if isinstance(weights, np.ndarray):
        weights = weights.tolist()
    result = [1.0] * total
    if isinstance(weights, (list, tuple)):
        if len(weights) != total:
            raise ValueError(f"Jumlah bobot ({len(weights)}) harus sama dengan jumlah soal ({total})")
        pairs = list(enumerate(weights, start=1))
    elif isinstance(weights, dict):
        pairs = [(int(str(k).strip()), v) for k, v in weights.items()]
    else:
        raise ValueError("weights harus berupa list, dict, atau None")
    for number, value in pairs:
        if not 1 <= number <= total:
            raise ValueError(f"Nomor soal {number} pada bobot di luar rentang 1-{total}")
        if isinstance(value, bool) or not isinstance(value, (int, float)) or value < 0 or value != value:
            raise ValueError(f"Bobot soal {number} harus angka >= 0 (ditemukan: {value!r})")
        result[number - 1] = float(value)
    return result


def grade_paper(student_answers, answer_key, weights=None):
    """Bandingkan jawaban siswa dengan kunci dan hitung skor berbobot.

    student_answers : list (indeks 0 = soal 1) atau dict {nomor: huruf}; None = kosong.
    answer_key      : list / dict / string huruf, mis. ["A", "C"] atau "AC". Panjang kunci
                      menentukan jumlah soal.
    weights         : None (semua bobot 1), list, atau dict bobot per soal.

    Jawaban kosong dan arsiran ganda (mis. "AB") tidak pernah dihitung benar.
    Mengembalikan dict: total_score, max_score, benar, salah, kosong, ganda, persentase,
    dan details (daftar per soal: nomor, jawaban, kunci, benar, bobot).
    `salah` = jumlah soal - benar (kosong dan ganda ikut dihitung salah), sama seperti
    database_manager.py.
    """
    total = _infer_total(answer_key)
    key = _normalize_answers(answer_key, total, "kunci jawaban", is_key=True)
    answers = _normalize_answers(student_answers, total, "jawaban siswa")
    wts = _normalize_weights(weights, total)

    score = 0.0
    correct = blank = multi = 0
    details = []
    for i in range(total):
        ans, k = answers[i], key[i]
        ok = ans == k
        if ok:
            score += wts[i]
            correct += 1
        if ans is None:
            blank += 1
        elif len(ans) > 1:
            multi += 1
        details.append({"nomor": i + 1, "jawaban": ans, "kunci": k, "benar": ok, "bobot": wts[i]})

    max_score = sum(wts)
    return {
        "total_score": round(score, 2),
        "max_score": round(max_score, 2),
        "benar": correct,
        "salah": total - correct,
        "kosong": blank,
        "ganda": multi,
        "persentase": round(score / max_score * 100, 2) if max_score > 0 else 0.0,
        "details": details,
    }


# --------------------------------------------------------------------------
# 4. Umpan balik visual
# --------------------------------------------------------------------------
def draw_visual_feedback(warped_image, question_coords, student_answers, answer_key,
                         show_correct_key=False):
    """Lingkari jawaban siswa: HIJAU bila benar, MERAH bila salah.

    warped_image     : kertas lurus (keluaran detect_and_warp()).
    question_coords  : keluaran get_bubble_coordinates().
    student_answers  : keluaran process_bubbles() (list atau dict).
    answer_key       : kunci jawaban (list, dict, atau string).
    show_correct_key : bila True, soal yang salah/kosong juga diberi lingkaran hijau tipis
                       pada kunci yang benar.

    Arsiran ganda ditandai merah pada setiap bulatan yang diarsir. Mengembalikan salinan
    gambar yang sudah ditandai; gambar asli tidak diubah.
    """
    canvas = _to_bgr(_canonical(warped_image))
    total = max(question_coords)
    answers = _normalize_answers(student_answers, total, "jawaban siswa")
    key = _normalize_answers(answer_key, total, "kunci jawaban", is_key=True)

    radius = int(round(BUBBLE_R_MM * SCALE * 1.35))
    thick = max(2, int(round(SCALE * 0.45)))
    for q, centers in question_coords.items():
        ans, correct = answers[q - 1], key[q - 1]
        if ans:
            color = GREEN if ans == correct else RED
            for letter in ans:
                if letter in OPTIONS:
                    cv2.circle(canvas, tuple(int(v) for v in centers[OPTIONS.index(letter)]),
                               radius, color, thick, cv2.LINE_AA)
        if show_correct_key and ans != correct:
            cv2.circle(canvas, tuple(int(v) for v in centers[OPTIONS.index(correct)]),
                       radius, GREEN, max(1, thick // 2), cv2.LINE_AA)
    return canvas


# --------------------------------------------------------------------------
# Satu panggilan untuk seluruh alur
# --------------------------------------------------------------------------
def scan_ljk(image_path, answer_key, total_questions=None, weights=None,
             feedback_path=None, show_correct_key=False, **bubble_options):
    """Foto LJK -> jawaban siswa -> skor -> gambar umpan balik.

    total_questions : bila None, dipakai panjang answer_key; jika QR code terbaca dan jumlah
                      soalnya berbeda dari kunci, ValueError dilempar (kunci / LJK salah).
    feedback_path   : (opsional) simpan gambar hasil tanda hijau/merah (.png / .jpg).
    **bubble_options: diteruskan ke process_bubbles() (fill_threshold, dst.).

    Melempar OMRError bila kotak jangkar tidak ditemukan atau tata letak tidak cocok.
    Mengembalikan dict: metadata (isi QR atau None), student_answers, fill_ratios,
    alignment_score, grade (hasil grade_paper), warped, feedback.
    """
    key_total = _infer_total(answer_key)
    warped = detect_and_warp(image_path, total_questions=total_questions or key_total)
    metadata = read_qr_metadata(warped)

    if total_questions is None:
        total_questions = key_total
        qr_total = metadata.get("Jumlah_Soal") if metadata else None
        if isinstance(qr_total, int) and qr_total != key_total:
            raise ValueError(
                f"Jumlah soal pada QR code ({qr_total}) berbeda dengan kunci jawaban "
                f"({key_total}). Kemungkinan kunci kuis salah atau LJK dari kuis lain."
            )

    score = _alignment_score(cv2.cvtColor(warped, cv2.COLOR_BGR2GRAY), total_questions)
    if score < ALIGN_MIN_SCORE:
        raise OMRError(
            f"Tata letak bulatan pada foto tidak cocok dengan LJK {total_questions} soal "
            f"(skor kecocokan {score:.2f}). Periksa jumlah soal, arah kertas, dan pastikan "
            "keempat kotak hitam sudut terlihat utuh."
        )

    answers, ratios = process_bubbles(warped, total_questions, **bubble_options)
    grade = grade_paper(answers, answer_key, weights)
    feedback = draw_visual_feedback(
        warped, get_bubble_coordinates(total_questions), answers, answer_key,
        show_correct_key=show_correct_key,
    )
    if feedback_path:
        ext = os.path.splitext(feedback_path)[1] or ".png"
        cv2.imencode(ext, feedback)[1].tofile(feedback_path)

    return {
        "metadata": metadata,
        "student_answers": answers,
        "fill_ratios": ratios,
        "alignment_score": round(score, 3),
        "grade": grade,
        "warped": warped,
        "feedback": feedback,
    }


# --------------------------------------------------------------------------
# Command line
# --------------------------------------------------------------------------
def _main():
    parser = argparse.ArgumentParser(description="Periksa foto LJK dengan OMR")
    parser.add_argument("image", help="foto LJK (.jpg / .png)")
    parser.add_argument("--key", required=True,
                        help="kunci jawaban, mis. ABCDE... atau A,B,C,D,E,...")
    parser.add_argument("--total", type=int, default=None, help="jumlah soal (default: dari QR / panjang kunci)")
    parser.add_argument("--out", default=None, help="simpan gambar umpan balik hijau/merah")
    parser.add_argument("--show-key", action="store_true", help="tandai juga kunci yang benar")
    args = parser.parse_args()

    key = [k.strip() for k in args.key.split(",")] if "," in args.key else list(args.key.strip())
    try:
        res = scan_ljk(args.image, key, total_questions=args.total,
                       feedback_path=args.out, show_correct_key=args.show_key)
    except (OMRError, ValueError, FileNotFoundError) as exc:
        print("GAGAL:", exc, file=sys.stderr)
        sys.exit(1)

    grade = {k: v for k, v in res["grade"].items() if k != "details"}
    print(json.dumps({
        "metadata": res["metadata"],
        "jawaban": "".join(a if a else "-" for a in res["student_answers"]),
        "nilai": grade,
    }, ensure_ascii=False, indent=2))
    if args.out:
        print("Gambar umpan balik disimpan:", args.out)


if __name__ == "__main__":
    _main()
