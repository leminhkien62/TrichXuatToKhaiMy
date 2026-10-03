#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
parse_7501.py — Trích xuất dữ liệu từ file PDF "CBP Form 7501 - Entry Summary"
(bản chính + Continuation Sheet) ra Excel.

Cách dùng:
    python parse_7501.py <thư_mục_chứa_pdf> [--out ket_qua.xlsx]

Ý tưởng chính (giống script tham khảo, nhưng áp dụng cho form 7501):
  - Dò định dạng (loại) form dựa trên toạ độ x,y của các ô tiêu đề bảng:
      * Loại 1: form cũ  -> cột 29 (HTSUS/AD-CVD No), 30 (Gross Weight/Manifest Qty),
                31 (Net Quantity), 33 (HTSUS/AD-CVD/IRC Rate), 34 (Duty and IR Tax)
      * Loại 2: form mới -> cột 33 (HTSUS/ADA-CVD No), 34 (Gross Weight/Manifest Qty),
                35 (Net Quantity), 37 (HTSUS/ADA-CVD/IRC Rate), 38 (Duty and I.R. Tax)
      * Loại 1.1: form mới (bố cục cột 33-38 giống Loại 2) nhưng dùng chữ
                "AD/CVD" (không có chữ "ADA/CVD") giống Loại 1 -> xem chi tiết
                phần "LOẠI 1.1" bên dưới.
      * Không xác định được -> "loại khác"
  - Từ toạ độ x của các tiêu đề trên, tính các mốc ranh giới (midpoint) để suy ra
    vùng x động cho từng cột: [mã HTS + mô tả] | [trọng lượng/số lượng] |
    [trị giá] | [thuế suất] | [tiền thuế].
  - Gom các "word" (từ PyMuPDF) theo toạ độ y thành từng dòng trong bảng, rồi
    gán từng dòng vào đúng cột theo vùng x ở trên.
  - Dòng "Invoice Number .../EURU-xxxxxx" hoặc "Totals for Invoice" dùng để xác
    định ranh giới (toạ độ y) giữa các invoice trong cùng 1 entry -> gán số
    invoice cho từng dòng hàng (line item) nằm phía trên nó.
  - Mỗi "Line No." (vd 001) có thể có NHIỀU dòng dữ liệu con (nhiều mã HTS,
    nhiều dòng trọng lượng/trị giá/thuế suất/tiền thuế — vd khi 1 line có
    nhiều HTS/AD-CVD khác nhau). Các dòng con này KHÔNG bị gộp/cộng dồn lại
    nữa — mỗi dòng con xuất ra thành 1 dòng riêng trong Excel, dùng chung
    "So_dong_Line_No" nhưng có thêm cột "So_dong_con" để phân biệt.
  - Kết quả xuất ra Excel: mỗi dòng = 1 dòng con dữ liệu của entry, kèm cột
    ghi rõ "Loại 1" / "Loại 2" / "Loại 1.1" / "Loại khác" để biết định dạng đã dò được.

  ─────────────────────────────────────────────────────────────────────────
  BỔ SUNG (mới): trích xuất thêm 15 TRƯỜNG HEADER đầu tiên của Entry Summary
  (mục 1 "Filer Code/Entry Number" đến mục 15 "Export Date"), áp dụng ĐƯỢC
  CHO CẢ loại 1, loại 2 LẪN loại 3 (form 5/22, 07/25 và bản 02/26), vì cả ba
  form đều dùng chung cách đánh số & tên trường cho các mục 1..15 (chỉ khác
  bố cục từ mục 16/21 trở đi và cách đặt cột hàng hoá 29-34/33-38). Cách dò
  cũng dùng kỹ thuật toạ độ (x,y) tương tự phần dò cột 29-34/33-38 ở trên:
  tìm toạ độ các ô nhãn "1." .. "15." trên trang đầu, suy ra vùng x/y động
  cho từng ô, rồi gom "word" nằm trong vùng đó làm giá trị. Xem hàm
  extract_header_1_15() bên dưới. Các trường này được gắn thêm vào MỖI dòng
  Excel (và vào sheet tổng hợp) để tiện lọc/à so sánh giữa loại 1, 2 và 3.
  ─────────────────────────────────────────────────────────────────────────

  ─────────────────────────────────────────────────────────────────────────
  BỔ SUNG (mới): LOẠI 1.1.
  Một số form 7501 bản mới (vd "CBP Form 7501 (02/26)") dùng LẠI cách đánh
  số cột hàng hoá GIỐNG HỆT Loại 2 (33. HTSUS No/AD-CVD No, 34. Gross
  Weight/Manifest Qty, 35. Net Quantity, 36. Entered Value, 37. HTSUS Rate/
  AD-CVD Rate/IRC Rate, 38. Duty and IR Tax) NHƯNG lại dùng chữ "AD/CVD"
  (KHÔNG có chữ "ADA/CVD") giống như Loại 1. Vì vậy không thể chỉ dựa vào có
  hay không có chữ "ADA/CVD" để phân biệt Loại 1 và Loại 1.1 (cả hai đều chỉ
  có "AD/CVD") — phải phân biệt thêm bằng bố cục cột: Loại 1.1 có nhãn "38."
  đứng ngay trước cụm "Duty and IR Tax" (vì cột tiền thuế là cột 38), trong
  khi Loại 1 có nhãn "34." đứng ngay trước cụm "Duty and IR Tax" (vì cột
  tiền thuế là cột 34). Xem cờ `uses_33_38_layout` trong detect_format().

  Về phần dò số Invoice ("Totals for Invoice"): Loại 1.1 hoạt động GIỐNG Loại 1
  — dòng "Totals for Invoice" (kèm số invoice ở dòng/khối NGAY SAU) nằm ở
  PHÍA SAU các dòng hàng (line item) mà nó tổng hợp, nên khi tra invoice cho
  1 dòng hàng phải tìm mốc invoice gần nhất có toạ độ y LỚN HƠN (tìm về phía
  sau) — khác với Loại 2 (dòng "Invoice Number .../EURU-xxxxxx" nằm TRƯỚC
  các dòng hàng của nó, nên phải tìm mốc invoice gần nhất có y NHỎ HƠN). Xem
  hàm invoice_for_y() bên dưới: Loại 1 và Loại 1.1 dùng chung 1 nhánh xử lý.
  ─────────────────────────────────────────────────────────────────────────
"""

import argparse
import re
import sys
from pathlib import Path

import pymupdf  # PyMuPDF
import pandas as pd

# ───────────────────────── Cấu hình cột theo từng loại form ─────────────────────────
# Với mỗi loại, khai báo text của các ô tiêu đề số thứ tự (vd "29.", "33.") cần tìm
# để lấy toạ độ x làm "tâm cột", từ đó suy ra ranh giới động giữa các cột.
FORMAT_DEFS = {
    "loai_1": {
        "field_labels": ["29.", "30.", "31.", "32.", "33.", "34."],
        "roles": ["ma_hts", "trong_luong_sl", "trong_luong_sl", "tri_gia", "thue_suat", "tien_thue"],
        # roles: cột nào gộp vào cột nào (30 & 31 gộp thành 1 vùng "trong_luong_sl")
        "distinguish_text": "AD/CVD",   # có trong text nhưng KHÔNG có "ADA/CVD"
        "must_not_have": "ADA/CVD",
    },
    "loai_2": {
        "field_labels": ["33.", "34.", "35.", "36.", "37.", "38."],
        "roles": ["ma_hts", "trong_luong_sl", "trong_luong_sl", "tri_gia", "thue_suat", "tien_thue"],
        "distinguish_text": "ADA/CVD",
        "must_not_have": None,
    },
    "loai_1_1": {
        # Bố cục cột (số thứ tự 33-38) GIỐNG Loại 2, nhưng dùng chữ "AD/CVD"
        # (không có "ADA/CVD") giống Loại 1. Xem ghi chú "LOẠI 1.1" ở đầu file.
        "field_labels": ["33.", "34.", "35.", "36.", "37.", "38."],
        "roles": ["ma_hts", "trong_luong_sl", "trong_luong_sl", "tri_gia", "thue_suat", "tien_thue"],
        "distinguish_text": "AD/CVD",
        "must_not_have": "ADA/CVD",
    },
}

# ───────── Danh sách 15 trường header đầu tiên (GIỐNG NHAU ở loại 1, 2 & 3) ─────────
# Thứ tự & tên mục lấy theo đúng CBP Form 7501 (cả bản 5/22, bản 07/25 lẫn bản 02/26):
#   1 Filer Code/Entry Number   6 Port Code            11 Import Date
#   2 Entry Type                7 Entry Date           12 B/L or AWB Number
#   3 Summary Date              8 Importing Carrier    13 Manufacturer ID
#   4 Surety Number             9 Mode of Transport    14 Exporting Country
#   5 Bond Type                 10 Country of Origin   15 Export Date
FIELD_NAMES_1_15 = [
    (1, "Filer_Code_Entry_Number"),
    (2, "Entry_Type"),
    (3, "Summary_Date"),
    (4, "Surety_Number"),
    (5, "Bond_Type"),
    (6, "Port_Code"),
    (7, "Entry_Date"),
    (8, "Importing_Carrier"),
    (9, "Mode_of_Transport"),
    (10, "Country_of_Origin"),
    (11, "Import_Date"),
    (12, "BL_or_AWB_Number"),
    (13, "Manufacturer_ID"),
    (14, "Exporting_Country"),
    (15, "Export_Date"),
]

RE_HTS = re.compile(r"\b\d{4}\.\d{2}\.\d{2,4}\b")
RE_MONEY = re.compile(r"\$?\-?[\d,]+\.\d{2}")
RE_RATE = re.compile(r"\bFREE\b|\bFree\b|\d+(?:\.\d+)?%")
RE_LINE_START = re.compile(r"^(\d{3})\s+(.*)$")

# RE_INVOICE_NO = re.compile(r"Invoice\s+Number\s+(\S+/[A-Z0-9\-]+)", re.IGNORECASE)
# RE_TOTALS_INVOICE = re.compile(r"Totals?\s+for\s+Invoice", re.IGNORECASE)


RE_INVOICE_NO = re.compile(
    r"Invoice\s+Number\s+\S+/(\S+)",
    re.IGNORECASE
)
RE_TOTALS_INVOICE = re.compile(
    r"Totals?\s+for\s+Invoice\s+(\S+)",
    re.IGNORECASE
)

# Dùng để phân biệt Loại 1 (cột tiền thuế là "34.") với Loại 1.1 (cột tiền thuế
# là "38.") khi cả hai đều chỉ có chữ "AD/CVD" (không có "ADA/CVD").
RE_38_DUTY = re.compile(r"38\.\s*Duty\s+and\s+I\.?R\.?\s*Tax", re.IGNORECASE)

# Nhãn số thứ tự các ô header "1." .. "15." (dùng cho extract_header_1_15).
RE_FIELD_LABEL_1_15 = re.compile(r"^(\d{1,2})\.$")


def _words(page):
    """Lấy words đã sort theo (y, x)."""
    ws = page.get_text("words")  # (x0,y0,x1,y1,"word",block,line,word_no)
    ws.sort(key=lambda w: (round(w[1], 1), w[0]))
    return ws


def _group_rows(words, y_tol=3.0):
    """Gom words thành các dòng theo toạ độ y (word đã sort theo y,x)."""
    rows = []
    cur_y = None
    cur = []
    for w in words:
        y = w[1]
        if cur_y is None or abs(y - cur_y) <= y_tol:
            cur.append(w)
            cur_y = y if cur_y is None else cur_y
        else:
            rows.append((cur_y, cur))
            cur = [w]
            cur_y = y
    if cur:
        rows.append((cur_y, cur))
    return rows


def detect_format(all_words_flat, full_text):
    """
    Dò loại form (loai_1 / loai_2 / loai_1_1 / khac) + tính vùng x động của các
    cột dựa trên toạ độ x của các ô tiêu đề 29/30/31/33/34 (loại 1) hoặc
    33/34/35/37/38 (loại 2 và loại 3 — hai loại này dùng CHUNG cách đánh số
    cột, chỉ khác nhau ở chữ "ADA/CVD" so với "AD/CVD").
    Trả về (fmt_name, col_bounds) ; col_bounds là dict role -> (x_min, x_max)
    """
    has_ada = "ADA/CVD" in full_text
    has_adcvd_only = ("AD/CVD" in full_text) and not has_ada

    # Loại 1.1 dùng chung bố cục cột (33-38) với Loại 2, nên nhận biết bằng
    # việc nhãn "38." đứng ngay trước cụm "Duty and IR Tax" (cột tiền thuế
    # là cột 38). Loại 1 thì cột tiền thuế là cột "34." (không có "38.").
    uses_33_38_layout = bool(RE_38_DUTY.search(full_text))

    if has_ada:
        fmt = "loai_2"
    elif has_adcvd_only and uses_33_38_layout:
        fmt = "loai_1_1"
    elif has_adcvd_only:
        fmt = "loai_1"
    else:
        return "loai_khac", None

    defs = FORMAT_DEFS[fmt]
    labels = defs["field_labels"]

    # Tìm toạ độ x của từng nhãn số thứ tự ô "xx." — lấy occurrence đầu tiên
    # (giống nhau ở mọi trang vì header lặp lại trên continuation sheet).
    centers = {}
    for w in all_words_flat:
        text = w[4]
        for lab in labels:
            if lab not in centers and text == lab:
                centers[lab] = w[0]
        if len(centers) == len(labels):
            break

    if len(centers) < len(labels):
        # Không tìm đủ toạ độ mốc -> coi như "loại khác" để dò tiếp
        return "loai_khac", None

    xs = [centers[lab] for lab in labels]  # theo đúng thứ tự field
    # Ranh giới = trung điểm giữa 2 tâm liên tiếp
    b1 = (xs[0] + xs[1]) / 2
    b2 = (xs[2] + xs[3]) / 2   # ranh giới sau cột trọng lượng+số lượng (gộp 30&31 / 34&35)
    b3 = (xs[3] + xs[4]) / 2
    b4 = (xs[4] + xs[5]) / 2

    col_bounds = {
        "ma_hts_mota": (0, b1),
        "trong_luong_sl": (b1, b2),
        "tri_gia": (b2, b3),
        "thue_suat": (b3, b4),
        "tien_thue": (b4, 10000),
    }
    return fmt, col_bounds


def _col_of(x, col_bounds):
    for role, (lo, hi) in col_bounds.items():
        if lo <= x < hi:
            return role
    return "tien_thue"  # fallback: quá bên phải


def _table_header_bottom_y(words):
    """Toạ độ y ngay dưới header bảng (dòng có chữ 'Dollars') -> nơi bắt đầu dữ liệu."""
    for w in words:
        if w[4] == "Dollars":
            return w[1] + 6
    return None


def _table_end_y(words, header_y):
    """Toạ độ y kết thúc bảng trên 1 trang (trước 'Other Fee Summary' / cuối trang)."""
    candidates = []
    for w in words:
        if w[1] > (header_y or 0) and w[4] in ("Summary", "CBP") and w[1] < 750:
            candidates.append(w[1])
    if candidates:
        return min(candidates) - 2
    return max((w[1] for w in words), default=999) + 5


# ═══════════════════════════════════════════════════════════════════════
# ─── BỔ SUNG: trích xuất 15 trường header (mục 1 → 15) theo toạ độ ───────
# ═══════════════════════════════════════════════════════════════════════

def _header_words(page):
    """
    Lấy toàn bộ 'words' của TRANG ĐẦU (trang chứa bảng header mục 1..24 của
    Entry Summary), sort theo (y, x) để đọc đúng thứ tự trình bày trên form.
    Dùng riêng hàm này (thay vì _words) chỉ để rõ ràng về mục đích sử dụng,
    logic sort giống hệt _words().
    """
    ws = page.get_text("words")
    ws.sort(key=lambda w: (round(w[1], 1), w[0]))
    return ws


def _find_field_1_15_label_positions(words):
    """
    Dò toạ độ (x0, y0, y1) của các ô nhãn số thứ tự '1.' .. '15.' ở PHẦN
    HEADER (đầu trang 1) của form 7501 — áp dụng được cho CẢ loại 1, loại 2
    LẪN loại 3 vì cả ba form này dùng chung cách đánh số cho các trường
    1..15 (chỉ khác bố cục từ mục 16/21 trở đi).

    Chỉ lấy occurrence ĐẦU TIÊN của mỗi số 1..15 (occurrence trùng số ở
    bảng dòng hàng bên dưới chỉ xuất hiện từ mục 27 trở đi nên không đụng
    nhau với các nhãn 1..15 ở đây).

    Trả về dict: {so_thu_tu: (x0, y0, y1)}. Nếu dò được đủ 15 nhãn thì dừng
    sớm để tránh quét lố xuống các bảng khác trong trang.
    """
    labels = {}
    for w in words:
        m = RE_FIELD_LABEL_1_15.match(w[4])
        if m:
            n = int(m.group(1))
            if 1 <= n <= 15 and n not in labels:
                labels[n] = (w[0], w[1], w[3])
        if len(labels) == 15:
            break
    return labels


def extract_header_1_15(doc):
    """
    Trích xuất 15 trường đầu tiên của Entry Summary (Filer Code/Entry Number
    ... Export Date) dựa theo TOẠ ĐỘ — dùng chung được cho cả loại 1, loại 2
    và loại 3 vì thứ tự & tên trường mục 1..15 giống nhau ở cả ba định dạng:
        1 Filer Code/Entry Number   6 Port Code            11 Import Date
        2 Entry Type                7 Entry Date           12 B/L or AWB Number
        3 Summary Date              8 Importing Carrier    13 Manufacturer ID
        4 Surety Number             9 Mode of Transport    14 Exporting Country
        5 Bond Type                 10 Country of Origin   15 Export Date

    Cách làm (cùng kỹ thuật đã dùng để dò vùng x động của cột 29-34/33-38
    ở bảng dòng hàng, xem detect_format()):
      1) Tìm toạ độ (x, y0, y1) của các ô nhãn '1.' .. '15.' trên trang đầu
         (_find_field_1_15_label_positions).
      2) Các nhãn này (kèm caption, vd '1. Filer Code/Entry Number') nằm
         thành từng HÀNG NGANG (gom theo y, dung sai 3pt — giống _group_rows);
         GIÁ TRỊ thực tế nằm ở (các) HÀNG NGAY BÊN DƯỚI hàng nhãn/caption đó,
         thẳng cột theo x với nhãn tương ứng.
      3) Với mỗi nhãn trong 1 hàng, suy ra vùng x động [x_min, x_max) =
         [toạ độ x của chính nhãn đó, toạ độ x của nhãn liền sau CÙNG HÀNG)
         (nhãn cuối hàng thì lấy x + 200 làm biên phải tạm) — cùng nguyên lý
         "ranh giới = mốc x của tiêu đề kế tiếp" như đã dùng ở detect_format().
      4) Vùng y của giá trị = từ ngay dưới hàng nhãn hiện tại đến ngay trước
         hàng nhãn kế tiếp (hoặc +40pt nếu là hàng nhãn cuối cùng).
      5) Gom mọi 'word' rơi vào đúng vùng (x, y) của từng field làm giá trị,
         nối lại theo đúng thứ tự đọc (y rồi x).

    Nếu KHÔNG dò đủ 15 nhãn (vd trang lỗi, thiếu chữ, hoặc không phải trang
    Entry Summary) -> trả về dict rỗng {}, KHÔNG đoán mò để tránh gán sai
    giá trị.
    """
    page = doc[0]
    words = _header_words(page)
    labels = _find_field_1_15_label_positions(words)
    if len(labels) < 15:
        return {}

    # ---- Gom nhãn thành từng "hàng" theo toạ độ y (dùng lại nguyên lý của
    #      _group_rows, nhưng viết riêng vì input ở đây là dict, không phải
    #      list words) ----
    items = sorted(labels.items(), key=lambda kv: kv[0])  # theo số thứ tự 1..15
    y_tol = 3.0
    rows = []
    cur_row = []
    cur_y = None
    for n, (x, y0, y1) in items:
        if cur_y is None or abs(y0 - cur_y) <= y_tol:
            cur_row.append((n, x, y0, y1))
            cur_y = y0 if cur_y is None else cur_y
        else:
            rows.append(cur_row)
            cur_row = [(n, x, y0, y1)]
            cur_y = y0
    if cur_row:
        rows.append(cur_row)

    # ---- Tìm mốc y của "mục 16" (hoặc nhãn số nào đó ngay sau mục 15) để
    # làm biên dưới cho hàng nhãn CUỐI CÙNG (mục 12-15), tránh bị "tràn"
    # xuống lẫn vào các mục 16..24 nằm ngay bên dưới (các form có số hàng/số
    # cột mỗi hàng khác nhau nên không thể cố định 1 khoảng +Npt cứng). ----
    y_last_row_bottom = max(t[3] for t in rows[-1])
    RE_ANY_LABEL = re.compile(r"^(\d{1,2})\.$")
    next_section_ys = [
        w[1] for w in words
        if RE_ANY_LABEL.match(w[4]) and w[1] > y_last_row_bottom + 1
    ]
    y_next_section = min(next_section_ys) if next_section_ys else (y_last_row_bottom + 40)

    # ---- Với mỗi hàng nhãn, suy ra vùng x/y động cho từng field trong hàng ----
    field_regions = {}  # so_thu_tu -> (x_min, x_max, y_value_top, y_value_bottom)
    for ri, row in enumerate(rows):
        row_sorted = sorted(row, key=lambda t: t[1])  # theo x, trái sang phải
        xs = [t[1] for t in row_sorted]
        y_bottom_of_row = max(t[3] for t in row_sorted)
        if ri + 1 < len(rows):
            # Vùng giá trị kết thúc ngay trước hàng nhãn kế tiếp
            y_value_end = min(t[2] for t in rows[ri + 1]) - 1
        else:
            # Hàng nhãn cuối cùng -> kết thúc ngay trước mốc "mục 16..." kế
            # tiếp đã tìm ở trên (thay vì cộng cứng +40pt như trước, để
            # không lẫn dữ liệu của các mục 16-24 nằm ngay bên dưới).
            y_value_end = y_next_section - 1
        # Lưu ý: y1 (đáy) của bounding-box chữ nhãn thường LỚN HƠN 1 chút so
        # với toạ độ y thực tế của dòng giá trị ngay bên dưới (do khoảng hở
        # descender của font) — quan sát thực tế trên form 5/22 chênh lệch
        # chỉ ~1pt. Trừ thêm 3pt biên an toàn để không bỏ sót dòng giá trị
        # nằm sát ngay dưới hàng nhãn.
        y_top = y_bottom_of_row - 3
        for i, (n, x, y0, y1) in enumerate(row_sorted):
            x_min = x
            x_max = xs[i + 1] if i + 1 < len(xs) else x + 200
            field_regions[n] = (x_min, x_max, y_top, y_value_end)

    # ---- Gom các 'word' rơi vào đúng vùng (x,y) của từng field làm giá trị ----
    raw_values = {n: [] for n in field_regions}
    for w in words:
        wx, wy, wtext = w[0], w[1], w[4]
        for n, (x_min, x_max, y_top, y_bot) in field_regions.items():
            if x_min <= wx < x_max and y_top < wy <= y_bot:
                raw_values[n].append((wy, wx, wtext))
                break  # 1 word chỉ thuộc về 1 field

    result = {}
    for n, key in FIELD_NAMES_1_15:
        toks = sorted(raw_values.get(n, []))  # sort theo (y, x) để đọc đúng thứ tự
        val = " ".join(t[2] for t in toks).strip()
        result[key] = val or None
    return result


def parse_pdf(pdf_path: Path) -> dict:
    doc = pymupdf.open(pdf_path)
    all_words_flat = []
    full_text_parts = []
    for page in doc:
        all_words_flat.extend(page.get_text("words"))
        full_text_parts.append(page.get_text())
    full_text = re.sub(r"\s+", " ", " ".join(full_text_parts))

    fmt, col_bounds = detect_format(all_words_flat, full_text)

    m = re.search(r"\b([A-Z0-9]{3}-\d{7}-\d)\b", full_text)
    entry_number = m.group(1) if m else None

    # ---- Trích xuất 15 trường header (mục 1 -> 15) — dò ĐỘC LẬP với fmt
    # (loại 1/2/3/khác) vì cách bố trí mục 1..15 giống nhau ở cả 3 loại form,
    # nên vẫn cố lấy được ngay cả khi bảng dòng hàng (27-40) không dò được. ----
    header_1_15 = extract_header_1_15(doc)

    line_items = []
    invoice_events = []  # (y_global, page_no_new, page_no_old, kind, value) kind in {"invoice_no","totals"}

    current_item = None  # dict for the line item currently being built
    expect_invoice_code_next = False
    PAGE_H = 792.0  # toạ độ y "ảo" nối các trang lại để so sánh across pages

    if fmt == "loai_khac":
        doc.close()
        return {
            "source_file": pdf_path.name,
            "entry_number": entry_number,
            # "format": "loai_khac",
            "col_bounds": None,
            "header_1_15": header_1_15,
            "line_items": [],
        }

    for pno, page in enumerate(doc):
        words = _words(page)
        header_y = _table_header_bottom_y(words)
        if header_y is None:
            # Trang không có bảng (vd trang bìa không có dòng item) -> bỏ qua
            continue
        end_y = _table_end_y(words, header_y)
        rows = _group_rows([w for w in words if header_y <= w[1] <= end_y])

        for y, row_words in rows:
            row_words = sorted(row_words, key=lambda w: w[0])
            row_text = " ".join(w[4] for w in row_words)
            g_y = y + pno * PAGE_H  # toạ độ y "toàn cục" xuyên suốt nhiều trang

            # ---- Ranh giới invoice ----
            m_inv = RE_INVOICE_NO.search(row_text)
            if m_inv:
                invoice_events.append((g_y, m_inv.group(1)))
                if current_item:
                    line_items.append(current_item)
                    current_item = None
                expect_invoice_code_next = False
                continue
            if RE_TOTALS_INVOICE.search(row_text):
                # Loại 1 và Loại 1.1: dòng "Totals for Invoice" không kèm số
                # invoice trên cùng dòng -> số invoice + các dòng tổng giá
                # trị nằm ở (những) dòng NGAY SAU (vd "EURU-328297 28,030.00
                # ... USD"). Toàn bộ khối này KHÔNG thuộc về line item nào
                # -> bỏ qua cho tới khi gặp "Line No" mới.
                expect_invoice_code_next = True
                continue

            # ---- Dòng bắt đầu 1 Line No mới (vd "001 RECIP VN 20% DUTY") ----
            # (kiểm tra TRƯỚC khối "expect_invoice_code_next" để không bỏ sót
            # dòng Line No ngay sau khối tổng-hợp-invoice)
            m_line = RE_LINE_START.match(row_text)
            if m_line and int(m_line.group(1)) <= 200:
                expect_invoice_code_next = False
                if current_item:
                    line_items.append(current_item)
                current_item = {
                    "line_no": m_line.group(1),
                    "rows": [],   # danh sách dòng dữ liệu con (mỗi dòng con = 1 dòng Excel, có mô tả riêng)
                    # "pending_desc": mô tả đang gom (từ dòng "001 ..." hoặc các dòng mô tả
                    # tràn dòng) nhưng CHƯA gán cho dòng con nào — sẽ gán cho dòng con dữ
                    # liệu (có HTS/số liệu) kế tiếp xuất hiện, rồi được xoá đi.
                    "pending_desc": [m_line.group(2)],
                    "melt_country": None,
                    "mpf": None,
                    "hmf": None,
                    "y": g_y,
                }
                continue

            if expect_invoice_code_next:
                # Dòng thuộc khối tổng-hợp-invoice (Invoice Value/Exchange/Entered
                # Value...) -> KHÔNG thuộc về line item nào, chỉ lấy mã invoice.



                # mã số INVOICE NO
                # m_code = re.search(r"\b([A-Z]{2,6}-\d{4,})\b", row_text)

                m_invoice_no = re.search(r"\b([A-Z]{2,6}-?\d{4,}\S*)", row_text)
                if m_invoice_no:
                    invoice_events.append((g_y, m_invoice_no.group(1)))
                continue

            if current_item is None:
                continue  # rác trước dòng line item đầu tiên (vd text ở footer trang trước)

            # ---- Phí MPF / HMF ----
            if "499" in row_text and re.search(r"MERCHANDISE PROCESSING FEE|Merchandise Processing Fee", row_text, re.I):
                amt = RE_MONEY.findall(row_text)
                current_item["mpf"] = amt[-1] if amt else None
                continue
            if "501" in row_text and re.search(r"HARBOR MAINTENANCE FEE|Harbor Maintenance Fee", row_text, re.I):
                amt = RE_MONEY.findall(row_text)
                current_item["hmf"] = amt[-1] if amt else None
                continue

            # ---- Melted/Poured Country (loại 2 và loại 3) ----
            m_melt = re.search(r"Melted/Poured Country/Region:\s*(\S+)", row_text)
            if m_melt:
                current_item["melt_country"] = m_melt.group(1)
                continue

            # ---- Gán từng word của DÒNG NÀY vào đúng cột theo vùng x động ----
            # Dùng buffer riêng cho từng dòng rồi mới quyết định gộp vào đâu, để
            # tránh chữ mô tả tràn cột (vd "VEHICLES") bị hiểu nhầm thành số liệu.
            row_cols = {"ma_hts": [], "mo_ta": [], "trong_luong_sl": [], "tri_gia": [], "thue_suat": [], "tien_thue": []}
            row_has_hts = False
            row_has_numeric_weight = False
            for w in row_words:
                role = _col_of(w[0], col_bounds)
                token = w[4]
                if role == "ma_hts_mota":
                    if RE_HTS.match(token):
                        row_cols["ma_hts"].append(token)
                        row_has_hts = True
                    else:
                        row_cols["mo_ta"].append(token)
                elif role == "trong_luong_sl":
                    row_cols["trong_luong_sl"].append(token)
                    if re.search(r"\d", token):
                        row_has_numeric_weight = True
                elif role == "tri_gia":
                    row_cols["tri_gia"].append(token)
                elif role == "thue_suat":
                    row_cols["thue_suat"].append(token)
                elif role == "tien_thue":
                    row_cols["tien_thue"].append(token)

            if row_has_hts or row_has_numeric_weight:
                # Dòng dữ liệu thật (có mã HTS và/hoặc số cân nặng/số lượng).
                # QUAN TRỌNG: không gộp vào 1 line item nữa — mỗi dòng dữ liệu
                # (vd mỗi mã HTS/AD-CVD khác nhau trong cùng 1 Line No) được
                # lưu thành 1 phần tử riêng trong "rows", để sau này xuất ra
                # thành 1 dòng Excel riêng, không bị cộng/gộp chung. Mô tả
                # hàng hoá của dòng con này = mô tả đang gom dở (pending_desc,
                # từ dòng "001 ..." hoặc các dòng mô tả tràn dòng phía trước)
                # + chữ mô tả nằm ngay trên chính dòng số liệu này (nếu có).
                desc_tokens = current_item["pending_desc"] + row_cols["mo_ta"]
                current_item["rows"].append({
                    "ma_hts": row_cols["ma_hts"],
                    "mo_ta": desc_tokens,
                    "trong_luong_sl": row_cols["trong_luong_sl"],
                    "tri_gia": row_cols["tri_gia"],
                    "thue_suat": row_cols["thue_suat"],
                    "tien_thue": row_cols["tien_thue"],
                })
                current_item["pending_desc"] = []  # đã gán cho dòng con -> xoá để không lặp lại
            else:
                # Dòng không có mã HTS lẫn số cân nặng -> hoặc là mô tả tràn cột
                # (vd "OTH PRTS,ACCES,MOTOR VEHICLES"), hoặc là dòng "Relationship"
                # (vd "C $437", "C678", "Y"). Không đưa vào trị giá/thuế suất/tiền
                # thuế để tránh làm sai số liệu; chỉ gom vào "pending_desc" để
                # gán cho dòng con dữ liệu kế tiếp (mỗi dòng con có mô tả riêng).
                if not re.match(r"^C\d*(\s+\$?[\d,]+)?$|^Y$|^C\s+\$[\d,]+$", row_text.strip()):
                    current_item["pending_desc"].extend(
                        row_cols["mo_ta"] + row_cols["trong_luong_sl"]
                        + row_cols["tri_gia"] + row_cols["thue_suat"] + row_cols["tien_thue"])

        # hết trang: nếu current_item đang dở dang thì để nó tiếp tục sang trang sau
    if current_item:
        line_items.append(current_item)
    doc.close()

    # Với mỗi line item: nếu sau dòng dữ liệu con CUỐI CÙNG vẫn còn "pending_desc"
    # chưa gán (vd mô tả tràn dòng nằm ngay dưới dòng số liệu cuối, trước khi qua
    # Line No/Invoice kế tiếp) -> gộp nốt vào mô tả của dòng con cuối cùng đó.
    # Nếu Line No không dò được dòng con nào cả (rows rỗng) thì giữ nguyên
    # pending_desc để dùng làm mô tả fallback khi xuất Excel.
    for it in line_items:
        leftover = it.get("pending_desc", [])
        if leftover and it["rows"]:
            it["rows"][-1]["mo_ta"].extend(leftover)
            it["pending_desc"] = []

    # ---- Gán invoice cho từng line item dựa theo mốc y toàn cục ----
    invoice_events.sort(key=lambda e: e[0])

    def invoice_for_y(gy):
        if fmt in ("loai_1", "loai_1_1"):
            # Loại 1 & Loại 1.1: dòng "Totals for Invoice / <số invoice>" nằm
            # SAU các line item mà nó tổng hợp -> lấy mốc invoice gần nhất
            # PHÍA SAU (y lớn hơn).
            for ey, inv in invoice_events:
                if ey >= gy:
                    return inv
            return None
        else:
            # Loại 2: dòng "Invoice Number .../..." nằm TRƯỚC các line item của nó
            # -> lấy mốc invoice gần nhất PHÍA TRƯỚC (y nhỏ hơn).
            cur = None
            for ey, inv in invoice_events:
                if ey <= gy:
                    cur = inv
                else:
                    break
            return cur

    out_items = []
    for it in line_items:
        invoice_no = invoice_for_y(it["y"])

        # điều chỉnh lại cho đúng
        if invoice_no.startswith("EURO"):
            invoice_no_correct = (
                "EURU"
                + ("" if len(invoice_no) > 4 and invoice_no[4] == "-" else "-")
                + invoice_no[4:]
            )
        else:
            invoice_no_correct = (
                invoice_no[:4]
                + ("" if len(invoice_no) > 4 and invoice_no[4] == "-" else "-")
                + invoice_no[4:]
            )

        # Mỗi dòng con trong "rows" -> 1 dòng Excel riêng, MÔ TẢ RIÊNG, KHÔNG
        # gộp/cộng chung với các dòng con khác.
        sub_rows = it["rows"] if it["rows"] else [None]
        so_dong_con_total = len(sub_rows)
        for idx, sr in enumerate(sub_rows, start=1):

            if sr is not None:
                weight, qty = _split_weight_qty(" ".join(sr["trong_luong_sl"]))

                # rate = _all_dedup(RE_RATE, " ".join(sr["thue_suat"]))

                rate = _all_dedup(RE_RATE, " ".join(sr["thue_suat"]))

                duty = _sum_money(" ".join(sr["tien_thue"]))
                value = _sum_money(" ".join(sr["tri_gia"]))
                hts = ", ".join(sr["ma_hts"]) or None
                desc = _clean_desc(" ".join(sr["mo_ta"]))
                
                if not hts :
                    continue

                # đây là phí xăng dầu
                # if  hts == "8708.99.8180":
                #     continue
            # else:
            #     # Không dò được dòng con dữ liệu nào cho Line No này -> vẫn
            #     # xuất 1 dòng, dùng mô tả gom được (nếu có) làm fallback.
            #     weight = qty = rate = duty = value = hts = None
            #     desc = _clean_desc(" ".join(it.get("pending_desc", [])))

            out_items.append({
                "Line_No": it["line_no"],
                # "So_dong_con": f"{idx}/{so_dong_con_total}",
                "Invoice_No": invoice_no,

                "Invoice_No_Correct": invoice_no_correct,
                # Mô tả hàng hoá tách riêng cho từng dòng con (không còn để trống
                # ở các dòng con sau nữa).

                # "Mo_ta_hang_hoa": desc,
                "HTSUS_No/AD/CVD No": hts,
                # "Gross Weight": weight,
                # "Manifest Qty": qty,
                # "Entered_Value": value,
                "Rate": rate,
                "Duty_and_IR_Tax": duty,
                # Phí MPF/HMF và nước nấu chảy là thông tin chung của cả Line No
                # (không tách theo từng mã HTS) -> chỉ điền ở dòng con đầu tiên.
                # "Phi_MPF_499": it["mpf"] if idx == 1 else "",
                # "Phi_HMF_501": it["hmf"] if idx == 1 else "",
                # "Nuoc_nau_chay_Melt_Country": it["melt_country"] if idx == 1 else "",
            })

    return {
        "source_file": pdf_path.name,
        "entry_number": entry_number,
        # "format": fmt,
        "col_bounds": col_bounds,
        "header_1_15": header_1_15,
        "line_items": out_items,
    }


def _split_weight_qty(text):
    """Tách 'trọng lượng + số lượng' đã gộp thành 2 giá trị riêng dựa trên đơn vị đi kèm (KG / NO / X / PCS...)."""
    weights = [f"{m} KG" for m in re.findall(r"([\d,]+\.?\d*)\s*KG", text) if m not in ("0", "0.00")]
    m_q = re.findall(r"([\d,]+(?:\.\d+)?)\s*(NO|PCS|BX|DOZ|SET)\b", text)
    qtys = [f"{n} {u}" for n, u in m_q if n not in ("0.00", "0")]
    if not qtys and re.search(r"(?:^|\s)X(?:\s|$)", text):
        qtys = ["X"]
    return (", ".join(weights) or None), (", ".join(qtys) or None)


def _all_dedup(pattern, text):
    seen, out = set(), []
    for m in pattern.findall(text):
        if m not in seen:
            seen.add(m)
            out.append(m)
    return ", ".join(out) if out else None


def _sum_money(text):
    vals = RE_MONEY.findall(text)
    if not vals:
        return None
    total = sum(float(v.replace("$", "").replace(",", "")) for v in vals)
    return round(total, 2)


def _clean_desc(text):
    text = re.sub(r"\s+", " ", text).strip()
    text = re.sub(r"^(C\d*|Y)\s*", "", text)
    return text


LOAI_LABEL = {"loai_1": "Loại 1", "loai_2": "Loại 2", "loai_1_1": "Loại 1.1", "loai_khac": "Loại khác"}



def run_gui():
    """Giao diện Windows: chọn thư mục PDF và thư mục xuất Excel."""
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
    import threading

    FIXED_OUTPUT_NAME = "TongHop.xlsx"

    # ── Bảng màu ──────────────────────────────────────────────────────
    COLOR_BG = "#F4F6FB"
    COLOR_CARD = "#FFFFFF"
    COLOR_PRIMARY = "#2F5CFF"
    COLOR_PRIMARY_DARK = "#2347CC"
    COLOR_TEXT = "#1C1F2A"
    COLOR_SUBTEXT = "#6B7280"
    COLOR_BORDER = "#E3E6EE"
    COLOR_SUCCESS = "#1F9D55"

    root = tk.Tk()
    root.title("Trích xuất dữ liệu tờ khai hóa đơn Mỹ")
    root.geometry("820x480")
    root.minsize(760, 440)
    root.configure(bg=COLOR_BG)

    # ── Style ttk ─────────────────────────────────────────────────────
    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass

    style.configure("App.TFrame", background=COLOR_BG)
    style.configure("Card.TFrame", background=COLOR_CARD)
    style.configure(
        "Title.TLabel",
        background=COLOR_BG,
        foreground=COLOR_TEXT,
        font=("Segoe UI", 18, "bold"),
    )
    style.configure(
        "Subtitle.TLabel",
        background=COLOR_BG,
        foreground=COLOR_SUBTEXT,
        font=("Segoe UI", 10),
    )
    style.configure(
        "CardLabel.TLabel",
        background=COLOR_CARD,
        foreground=COLOR_TEXT,
        font=("Segoe UI", 10, "bold"),
    )
    style.configure(
        "CardHint.TLabel",
        background=COLOR_CARD,
        foreground=COLOR_SUBTEXT,
        font=("Segoe UI", 9),
    )
    style.configure(
        "Fixed.TLabel",
        background=COLOR_CARD,
        foreground=COLOR_PRIMARY_DARK,
        font=("Segoe UI", 10, "bold"),
    )
    style.configure(
        "Status.TLabel",
        background=COLOR_BG,
        foreground=COLOR_TEXT,
        font=("Segoe UI", 10),
    )
    style.configure(
        "App.TEntry",
        fieldbackground="#FFFFFF",
        bordercolor=COLOR_BORDER,
        relief="flat",
        padding=8,
    )
    style.configure(
        "Secondary.TButton",
        background="#EEF1F8",
        foreground=COLOR_TEXT,
        font=("Segoe UI", 9, "bold"),
        padding=(14, 8),
        relief="flat",
        borderwidth=0,
    )
    style.map(
        "Secondary.TButton",
        background=[("active", "#E1E6F5")],
    )
    style.configure(
        "Primary.TButton",
        background=COLOR_PRIMARY,
        foreground="#FFFFFF",
        font=("Segoe UI", 11, "bold"),
        padding=(22, 12),
        relief="flat",
        borderwidth=0,
    )
    style.map(
        "Primary.TButton",
        background=[("active", COLOR_PRIMARY_DARK), ("disabled", "#AAB4D4")],
    )
    style.configure(
        "App.Horizontal.TProgressbar",
        troughcolor="#EEF1F8",
        background=COLOR_PRIMARY,
        bordercolor="#EEF1F8",
        lightcolor=COLOR_PRIMARY,
        darkcolor=COLOR_PRIMARY,
        thickness=12,
    )

    input_var = tk.StringVar()
    output_var = tk.StringVar()
    status_var = tk.StringVar(value="Sẵn sàng. Hãy chọn thư mục chứa PDF để bắt đầu.")

    # ── Bố cục tổng ───────────────────────────────────────────────────
    outer = ttk.Frame(root, style="App.TFrame", padding=24)
    outer.pack(fill="both", expand=True)
    outer.columnconfigure(0, weight=1)

    # Header
    header = ttk.Frame(outer, style="App.TFrame")
    header.grid(row=0, column=0, sticky="ew", pady=(0, 20))
    ttk.Label(
        header, text="📄  ENTRY SUMMARY — Trích xuất dữ liệu", style="Title.TLabel"
    ).pack(anchor="w")
    ttk.Label(
        header,
        text="Chọn thư mục chứa PDF và thư mục lưu kết quả, sau đó bấm Bắt đầu.",
        style="Subtitle.TLabel",
    ).pack(anchor="w", pady=(4, 0))

    # Card chứa form nhập liệu
    card = tk.Frame(
        outer, bg=COLOR_CARD, highlightbackground=COLOR_BORDER,
        highlightthickness=1, bd=0,
    )
    card.grid(row=1, column=0, sticky="ew")
    outer.rowconfigure(1, weight=0)
    card_inner = ttk.Frame(card, style="Card.TFrame", padding=22)
    card_inner.pack(fill="both", expand=True)
    card_inner.columnconfigure(0, weight=1)

    # -- Dòng 1: thư mục PDF
    ttk.Label(card_inner, text="1.  Thư mục chứa PDF", style="CardLabel.TLabel").grid(
        row=0, column=0, sticky="w"
    )
    row1 = ttk.Frame(card_inner, style="Card.TFrame")
    row1.grid(row=1, column=0, sticky="ew", pady=(6, 18))
    row1.columnconfigure(0, weight=1)
    input_entry = ttk.Entry(row1, textvariable=input_var, style="App.TEntry")
    input_entry.grid(row=0, column=0, sticky="ew", ipady=4)
    ttk.Button(
        row1, text="Chọn thư mục...", style="Secondary.TButton",
        command=lambda: choose_input(),
    ).grid(row=0, column=1, padx=(10, 0))

    # -- Dòng 2: thư mục xuất
    ttk.Label(card_inner, text="2.  Thư mục lưu file xuất", style="CardLabel.TLabel").grid(
        row=2, column=0, sticky="w"
    )
    row2 = ttk.Frame(card_inner, style="Card.TFrame")
    row2.grid(row=3, column=0, sticky="ew", pady=(6, 18))
    row2.columnconfigure(0, weight=1)
    output_entry = ttk.Entry(row2, textvariable=output_var, style="App.TEntry")
    output_entry.grid(row=0, column=0, sticky="ew", ipady=4)
    ttk.Button(
        row2, text="Chọn thư mục...", style="Secondary.TButton",
        command=lambda: choose_output(),
    ).grid(row=0, column=1, padx=(10, 0))

    # -- Dòng 3: tên file cố định
    fixed_row = ttk.Frame(card_inner, style="Card.TFrame")
    fixed_row.grid(row=4, column=0, sticky="w", pady=(0, 4))
    ttk.Label(fixed_row, text="Tên file xuất:", style="CardHint.TLabel").pack(side="left")
    ttk.Label(fixed_row, text=f"  {FIXED_OUTPUT_NAME}", style="Fixed.TLabel").pack(side="left")

    # Progress + trạng thái
    progress_frame = ttk.Frame(outer, style="App.TFrame")
    progress_frame.grid(row=2, column=0, sticky="ew", pady=(22, 0))
    progress_frame.columnconfigure(0, weight=1)

    progress = ttk.Progressbar(
        progress_frame, mode="determinate", style="App.Horizontal.TProgressbar"
    )
    progress.grid(row=0, column=0, sticky="ew")

    status_label = ttk.Label(progress_frame, textvariable=status_var, style="Status.TLabel")
    status_label.grid(row=1, column=0, sticky="w", pady=(10, 0))

    # Nút bắt đầu
    button_row = ttk.Frame(outer, style="App.TFrame")
    button_row.grid(row=3, column=0, sticky="ew", pady=(20, 0))
    button_row.columnconfigure(0, weight=1)

    run_button = ttk.Button(button_row, text="▶  Bắt đầu trích xuất", style="Primary.TButton")
    run_button.grid(row=0, column=1, sticky="e")

    def choose_input():
        path = filedialog.askdirectory(title="Chọn thư mục chứa file PDF")
        if path:
            input_var.set(path)
            # Mặc định lưu cùng thư mục PDF nếu chưa chọn thư mục xuất.
            if not output_var.get().strip():
                output_var.set(path)

    def choose_output():
        path = filedialog.askdirectory(title="Chọn thư mục lưu file Excel")
        if path:
            output_var.set(path)

    def set_enabled(enabled):
        state = "normal" if enabled else "disabled"
        run_button.config(state=state)

    def worker():
        try:
            input_dir = Path(input_var.get().strip())
            output_dir = Path(output_var.get().strip())

            if not input_dir.is_dir():
                raise ValueError("Thư mục chứa PDF không hợp lệ.")

            output_dir.mkdir(parents=True, exist_ok=True)
            out_path = output_dir / FIXED_OUTPUT_NAME

            pdf_files = sorted(
                p for p in input_dir.rglob("*")
                if p.is_file() and p.suffix.lower() == ".pdf"
            )

            if not pdf_files:
                raise ValueError(
                    f"Không tìm thấy file PDF nào trong:\n{input_dir}"
                )

            total = len(pdf_files)
            root.after(0, lambda: progress.config(maximum=total, value=0))

            all_rows = []
            summary_rows = []
            col_debug_rows = []
            header_rows = []

            for index, pdf_path in enumerate(pdf_files, start=1):
                root.after(
                    0,
                    status_var.set,
                    f"Đang xử lý {index}/{total}: {pdf_path.name}",
                )
                root.after(0, lambda v=index: progress.config(value=v))

                try:
                    result = parse_pdf(pdf_path)
                except Exception as e:
                    summary_rows.append({
                        "File": pdf_path.name,
                        "Entry_Number": None,
                        # "Loai": f"LỖI: {e}",
                        "So_dong_Line_Item": 0,
                    })
                    continue

                # loai_label = LOAI_LABEL.get(
                #     result["format"], result["format"]
                # )

                so_line_no_khac_nhau = len({
                    row["Line_No"] for row in result["line_items"]
                })

                summary_rows.append({
                    "File": result["source_file"],
                    "Entry_Number": result["entry_number"],
                    # "Loai": loai_label,
                    "So_Line_No": so_line_no_khac_nhau,
                    "So_dong_Excel": len(result["line_items"]),
                })

                header_row = {
                    "File": result["source_file"],
                    "Entry_Number": result["entry_number"],
                    # "Loai": loai_label,
                }
                header_row.update(result.get("header_1_15") or {})
                header_rows.append(header_row)

                for item in result["line_items"]:
                    row = {
                        "File": result["source_file"],
                        "Entry_Number": result["entry_number"],
                        # "Loai": loai_label,
                    }
                    row.update(result.get("header_1_15") or {})
                    row.update(item)
                    all_rows.append(row)

                if result["col_bounds"]:
                    for role, (lo, hi) in result["col_bounds"].items():
                        col_debug_rows.append({
                            "File": result["source_file"],
                            # "Loai": loai_label,
                            "Cot": role,
                            "X_min": round(lo, 1),
                            "X_max": round(hi, 1),
                        })

            df_items = pd.DataFrame(all_rows)

            # Giữ đúng cấu trúc Excel hiện tại của script.
            with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
                if not df_items.empty:
                    df_items.to_excel(
                        writer,
                        sheet_name="Chi_tiet_line_items",
                        index=False,
                    )
                else:
                    pd.DataFrame([{
                        "Ghi_chu": (
                            "Không trích được dòng dữ liệu nào từ "
                            "các PDF trong thư mục này"
                        )
                    }]).to_excel(
                        writer,
                        sheet_name="Chi_tiet_line_items",
                        index=False,
                    )

            root.after(0, set_enabled, True)
            root.after(
                0,
                status_var.set,
                f"✓ Hoàn tất. Đã xử lý {total} file PDF.",
            )
            root.after(
                0,
                lambda: messagebox.showinfo(
                    "Hoàn tất",
                    f"Đã xuất file:\n{out_path}",
                ),
            )

        except Exception as e:
            root.after(0, set_enabled, True)
            root.after(0, status_var.set, "Có lỗi xảy ra.")
            root.after(
                0,
                lambda err=str(e): messagebox.showerror(
                    "Lỗi",
                    err,
                ),
            )

    def start():
        if not input_var.get().strip():
            messagebox.showwarning(
                "Thiếu thông tin",
                "Hãy chọn thư mục chứa PDF.",
            )
            return

        if not output_var.get().strip():
            messagebox.showwarning(
                "Thiếu thông tin",
                "Hãy chọn thư mục lưu file xuất.",
            )
            return

        set_enabled(False)
        progress.config(value=0)
        status_var.set("Đang chuẩn bị...")

        threading.Thread(target=worker, daemon=True).start()

    run_button.config(command=start)

    root.mainloop()


if __name__ == "__main__":
    run_gui()