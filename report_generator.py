# -*- coding: utf-8 -*-
"""
report_generator.py
============================================================
merged_loan_report.xlsx থেকে বিভিন্ন ধরনের A4-Landscape ব্যাংক PDF রিপোর্ট বানানোর মডিউল।

সাপোর্টেড রিপোর্ট:
    1. Overdue Loan (নির্দিষ্ট তারিখ পর্যন্ত, oldest→newest) -- ঐচ্ছিক এক/একাধিক
       Union এবং (Union-এর ভেতরে) ঐচ্ছিক Village filter সহ
    2. Expired Loan List up to <date>    -- overdue_date <= date (তারিখসহ), oldest→newest
    3. Rescheduled Loan up to <date>     -- overdue_date > date এবং Reschedule No. > 0
    4. Union/Village-ভিত্তিক গ্রুপড রিপোর্ট (সংশোধিত Excel থেকে), প্রতি Union-এর শেষে
       সাবটোটাল: Total Loans, Total Balance, Total Due, Total Overdue (রেফারেন্স তারিখ
       পর্যন্ত overdue-এর সংখ্যা ও ব্যালেন্স), Total Rescheduled (রেফারেন্স তারিখের পরের
       overdue + Reschedule No. > 0 রো-গুলোর সংখ্যা ও ব্যালেন্স) -- এবং সব Union শেষে
       Grand Total (সব সাবটোটালের যোগফল)।
    5. Due Amount Report

    প্রতিটা রিপোর্টে যেখানেই Union বাছাইয়ের অপশন আছে, সেখানে ঐচ্ছিকভাবে সেই
    Union(গুলো)-র ভেতরের Village-ও বাছাই করা যায় (খালি রাখলে সব Village আসবে)।

    প্রতিটা রিপোর্টের টেবিলে "Loan Case"-এর বামে একটা "Sl." (Serial) কলাম থাকে।
    সাধারণ রিপোর্টে এটা টানা ১, ২, ৩... — Union/Village রিপোর্টে প্রতিটা Union-এর
    জন্য আলাদাভাবে ১ থেকে শুরু হয়।

প্রতিটা PDF-এর প্রতি পেজেই উপরে ব্যাংকের লোগো + নাম + শাখা + রিপোর্ট-টাইটেল, এবং
ডান পাশে Print Date/Time + "@Md. Noushad Ahmed" থাকে।
"""
import os
import re
from collections import defaultdict
from datetime import date, datetime, timezone, timedelta

import openpyxl
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import (Paragraph, SimpleDocTemplate, Spacer, Table,
                                 TableStyle)

BANK_NAME_DEFAULT = "KARMASANGSTHAN BANK"
FONT_REGULAR = "Helvetica"
FONT_BOLD = "Helvetica-Bold"
BD_TZ = timezone(timedelta(hours=6))  # Bangladesh Standard Time, GMT+6 -- সার্ভারের timezone
                                       # যাই হোক না কেন, Print Date/Time সবসময় এই timezone-এ দেখানো হয়

# Excel-এর কলাম অর্ডারের সাথে হুবহু মিল রেখে (pdf_processor.write_excel দেখুন)
COLUMNS = [
    ("prefixed_loan_case", "Loan"),
    ("borrower", "Borrower"),
    ("father", "Father"),
    ("spouse", "Spouse"),
    ("village", "Village"),
    ("union", "Union"),
    ("phone", "Phone"),
    ("overdue_date", "Overdue"),
    ("installment", "Installment"),
    ("bal_principal", "Principal"),
    ("bal_interest", "Interest"),
    ("bal_total", "Balance"),
    ("due_amount", "Due"),
    ("reschedule_no", "Res."),
    ("blank_col", "Comment"),
]

_AMOUNT_KEYS = {"installment", "bal_principal", "bal_interest", "bal_total", "due_amount"}

# কলামগুলোর আপেক্ষিক প্রস্থ -- সংখ্যা/তারিখ কলাম সরু, নাম/ঠিকানা কলাম চওড়া
_COLUMN_WEIGHTS = {
    "prefixed_loan_case": 0.8,
    "borrower": 1.3,
    "father": 1.3,
    "spouse": 1.3,
    "village": 1.0,
    "union": 1.0,
    "phone": 1.15,
    "overdue_date": 0.85,
    "installment": 0.95,
    "bal_principal": 0.8,
    "bal_interest": 0.7,
    "bal_total": 0.8,
    "due_amount": 0.65,
    "reschedule_no": 0.55,
    "blank_col": 0.8,
}


# =============================================================================
# ১. merged Excel পড়া
# =============================================================================
def read_merged_excel(path):
    """merged_loan_report.xlsx (2-স্তরের হেডার, ৩ নং রো থেকে ডেটা) পড়ে dict-এর লিস্ট রিটার্ন করে।"""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    keys = [k for k, _ in COLUMNS]
    rows = []
    for r in ws.iter_rows(min_row=3, values_only=True):
        if r is None or all(v is None or str(v).strip() == "" for v in r[:2]):
            continue
        d = {k: r[i] if i < len(r) else None for i, k in enumerate(keys)}
        rows.append(d)
    return rows


def parse_ddmmyyyy(value):
    """'dd/mm/yyyy' স্ট্রিং বা datetime/date অবজেক্টকে date-এ রূপান্তর করে; না পারলে None।"""
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    s = str(value).strip()
    if not s:
        return None
    for fmt in ("%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def _num(v):
    try:
        if v is None or str(v).strip() == "":
            return 0.0
        return float(str(v).replace(",", ""))
    except (TypeError, ValueError):
        return 0.0


# =============================================================================
# ২. ফিল্টার + সর্ট ফাংশনসমূহ
# =============================================================================
def _apply_union_filter(rows, unions):
    """unions দেওয়া থাকলে শুধু সেই Union(গুলো)-র রো রাখে; না থাকলে সব রো ফেরত দেয়।"""
    union_set = {u.strip().lower() for u in (unions or []) if u and u.strip()}
    if not union_set:
        return rows
    return [d for d in rows if (d.get("union") or "").strip().lower() in union_set]


def _apply_village_filter(rows, villages):
    """villages দেওয়া থাকলে শুধু সেই Village(গুলো)-র রো রাখে; না থাকলে সব রো ফেরত দেয়।"""
    village_set = {v.strip().lower() for v in (villages or []) if v and v.strip()}
    if not village_set:
        return rows
    return [d for d in rows if (d.get("village") or "").strip().lower() in village_set]


def _apply_union_village_map_filter(rows, union_village_map):
    """union_village_map: {union_name: [village, ...]} অনুযায়ী ফিল্টার করে।
    - union_village_map খালি/None হলে -> কোনো ফিল্টার নেই, সব রো ফেরত।
    - কোনো Union-এর জন্য villages লিস্ট খালি থাকলে -> সেই Union-এর সব Village।
    - villages লিস্ট থাকলে -> শুধু সেই Village(গুলো)।
    প্রতিটা Union-এর village selection সম্পূর্ণ স্বতন্ত্র — একটা Union-এ village বাছাই
    করলে অন্য (বাছাইকৃত) Union-এর রো-কে প্রভাবিত করে না।"""
    if not union_village_map:
        return rows
    norm_map = {}
    for u, vills in union_village_map.items():
        u_key = (u or "").strip().lower()
        if not u_key:
            continue
        norm_map[u_key] = {v.strip().lower() for v in (vills or []) if v and v.strip()}
    if not norm_map:
        return rows
    out = []
    for d in rows:
        u_key = (d.get("union") or "").strip().lower()
        if u_key not in norm_map:
            continue
        v_set = norm_map[u_key]
        if not v_set or (d.get("village") or "").strip().lower() in v_set:
            out.append(d)
    return out


def describe_union_village_selection(union_village_map):
    """UI/টাইটেলে দেখানোর জন্য মানুষ-পড়ার-যোগ্য বর্ণনা বানায়, যেমন:
    'Joynagar (Partial), Kashipur' -- village নামগুলো সরাসরি না দেখিয়ে, যে Union-এ
    village বাছাই করা হয়েছে (অর্থাৎ পুরো Union নয়) তার পাশে শুধু '(Partial)' বসে।
    Village বাছাই না থাকা Union-এর নামের পাশে কোনো চিহ্ন থাকে না। union_village_map
    খালি হলে খালি string ফেরত দেয় (মানে সব Union, সব Village)।"""
    if not union_village_map:
        return ""
    parts = []
    for u in sorted(union_village_map.keys()):
        vills = union_village_map.get(u) or []
        if vills:
            parts.append(f"{u} (Partial)")
        else:
            parts.append(u)
    return ", ".join(parts)


def get_unions(rows):
    """সব রো থেকে ইউনিক, sorted Union-এর তালিকা।"""
    return sorted({(d.get("union") or "").strip() for d in rows if d.get("union")})


def get_villages_for_unions(rows, unions=None):
    """দেওয়া Union(গুলো)-র মধ্যে থাকা ইউনিক, sorted Village-এর তালিকা।
    unions খালি থাকলে সব রো-এর Village ফেরত দেয়।"""
    base = _apply_union_filter(rows, unions)
    return sorted({(d.get("village") or "").strip() for d in base if d.get("village")})


def filter_overdue(rows, start_date, end_date, union_village_map=None):
    """start_date <= overdue_date <= end_date (দুই পাশই inclusive) -- ঐচ্ছিক
    per-Union Village filter সহ -- oldest → newest sort।"""
    out = [
        d for d in rows
        if (dt := parse_ddmmyyyy(d.get("overdue_date"))) and start_date <= dt <= end_date
    ]
    out = _apply_union_village_map_filter(out, union_village_map)
    out.sort(key=lambda d: parse_ddmmyyyy(d.get("overdue_date")))
    return out


def filter_expired(rows, before, union_village_map=None):
    """overdue_date <= before (তারিখসহ, inclusive) -- ঐচ্ছিক per-Union Village filter সহ --
    Union অনুযায়ী সাজানো, প্রতি Union-এর ভেতরে oldest → newest।"""
    out = [d for d in rows if (dt := parse_ddmmyyyy(d.get("overdue_date"))) and dt <= before]
    out = _apply_union_village_map_filter(out, union_village_map)
    out.sort(key=lambda d: ((d.get("union") or "").strip().lower(), parse_ddmmyyyy(d.get("overdue_date"))))
    return out


def filter_rescheduled(rows, after, union_village_map=None):
    """overdue_date > after এবং Reschedule No. > 0 -- ঐচ্ছিক per-Union Village filter সহ --
    Union অনুযায়ী সাজানো, প্রতি Union-এর ভেতরে oldest → newest।"""
    out = [
        d for d in rows
        if (dt := parse_ddmmyyyy(d.get("overdue_date"))) and dt > after and _num(d.get("reschedule_no")) > 0
    ]
    out = _apply_union_village_map_filter(out, union_village_map)
    out.sort(key=lambda d: ((d.get("union") or "").strip().lower(), parse_ddmmyyyy(d.get("overdue_date"))))
    return out


def filter_due_amount(rows, union_village_map=None):
    """যেসব রো-তে Due Amount উপলব্ধ (> 0), শুধু সেগুলো -- ঐচ্ছিক per-Union Village filter সহ --
    Union অনুযায়ী সাজানো, প্রতি Union-এর ভেতরে oldest → newest।"""
    out = [d for d in rows if _num(d.get("due_amount")) > 0]
    out = _apply_union_village_map_filter(out, union_village_map)
    out.sort(key=lambda d: ((d.get("union") or "").strip().lower(),
                             parse_ddmmyyyy(d.get("overdue_date")) or date.min))
    return out


def group_by_union_village(rows, union_village_map=None, ref_date=None):
    """Union অনুযায়ী গ্রুপ (Union নাম অনুযায়ী sorted), প্রতি গ্রুপে Village অনুযায়ী sort।
    ঐচ্ছিক per-Union Village filter সহ। প্রতিটা Union-গ্রুপের সাথে সাবটোটাল থাকে:
        count           -- মোট লোন সংখ্যা
        balance         -- মোট Balance (bal_total যোগফল)
        due             -- মোট Due Amount যোগফল
        overdue_count   -- overdue_date <= ref_date এমন রো-র সংখ্যা
        overdue_balance -- ঐ রো-গুলোর Balance যোগফল
        resch_count     -- overdue_date > ref_date এবং Reschedule No. > 0 এমন রো-র সংখ্যা
        resch_balance   -- ঐ রো-গুলোর Balance যোগফল
    ref_date না দিলে আজকের তারিখ ধরা হয়।
    রিটার্ন: [(union_name, [row, ...], subtotal_dict), ...]"""
    if ref_date is None:
        ref_date = date.today()
    rows = _apply_union_village_map_filter(rows, union_village_map)
    groups = defaultdict(list)
    for d in rows:
        union = (d.get("union") or "Unknown").strip() or "Unknown"
        groups[union].append(d)

    result = []
    for union in sorted(groups.keys()):
        group_rows = sorted(groups[union], key=lambda d: (d.get("village") or "").strip())

        overdue_rows = [
            d for d in group_rows
            if (dt := parse_ddmmyyyy(d.get("overdue_date"))) and dt <= ref_date
        ]
        resch_rows = [
            d for d in group_rows
            if (dt := parse_ddmmyyyy(d.get("overdue_date"))) and dt > ref_date
            and _num(d.get("reschedule_no")) > 0
        ]

        subtotal = {
            "count": len(group_rows),
            "balance": sum(_num(d.get("bal_total")) for d in group_rows),
            "due": sum(_num(d.get("due_amount")) for d in group_rows),
            "overdue_count": len(overdue_rows),
            "overdue_balance": sum(_num(d.get("bal_total")) for d in overdue_rows),
            "resch_count": len(resch_rows),
            "resch_balance": sum(_num(d.get("bal_total")) for d in resch_rows),
        }
        result.append((union, group_rows, subtotal))
    return result


def build_output_filename(report_key, union_village_map=None, start=None, end=None, single_date=None):
    """PDF ফাইলের নাম dynamic ভাবে বানায়, যেমন:
    Overdue_Joynagar-Partial_Kashipur_15-03-2025_to_30-06-2026.pdf
    Expired_AllUnion_29-08-2026.pdf
    """
    def _clean(s):
        return re.sub(r"[^A-Za-z0-9]+", "", s) or "Union"

    parts = [report_key]
    if union_village_map:
        union_parts = []
        for u in sorted(union_village_map.keys()):
            vills = union_village_map.get(u) or []
            if vills:
                union_parts.append(_clean(u) + "-Partial")
            else:
                union_parts.append(_clean(u))
        parts.append("_".join(union_parts))
    else:
        parts.append("AllUnion")
    if start and end:
        parts.append(f"{start.strftime('%d-%m-%Y')}_to_{end.strftime('%d-%m-%Y')}")
    elif single_date:
        parts.append(single_date.strftime("%d-%m-%Y"))
    return "_".join(parts) + ".pdf"


# =============================================================================
# ৩. PDF জেনারেশন (A4 Landscape, প্রতি পেজে ব্যাংক-হেডার)
# =============================================================================
def _draw_header(canvas, doc, bank_name, branch_name, logo_path, title_text, print_dt_str=None):
    canvas.saveState()
    page_w, page_h = landscape(A4)

    subtitle = "A State Owned Financial Institution"
    branch_line = f"{branch_name}"

    bank_font, bank_size = FONT_BOLD, 13
    sub_font, sub_size = FONT_REGULAR, 8
    branch_font, branch_size = FONT_REGULAR, 9

    # টেক্সট তিন লাইনের y-position (আগের মতোই)
    y_bank = page_h - 20 * mm
    y_sub = page_h - 25 * mm
    y_branch = page_h - 30 * mm

    # টেক্সট-ব্লকের ভার্টিক্যাল সেন্টার (প্রথম আর শেষ লাইনের মাঝামাঝি)
    text_block_center_y = (y_bank + y_branch) / 2

    widths = [
        canvas.stringWidth(bank_name, bank_font, bank_size),
        canvas.stringWidth(subtitle, sub_font, sub_size),
        canvas.stringWidth(branch_line, branch_font, branch_size),
    ]
    max_text_width = max(widths)
    text_left_edge = page_w / 2 - max_text_width / 2

    logo_size = 16 * mm
    gap = 3 * mm
    logo_x = text_left_edge - gap - logo_size
    # লোগোর কেন্দ্র = text_block_center_y -> bottom-y = center - size/2
    logo_y = text_block_center_y - logo_size / 2

    if logo_path and os.path.exists(logo_path):
        canvas.drawImage(
            logo_path, logo_x, logo_y,
            width=logo_size, height=logo_size, mask="auto", preserveAspectRatio=True,
        )

    canvas.setFont(bank_font, bank_size)
    canvas.drawCentredString(page_w / 2, y_bank, bank_name)
    canvas.setFont(sub_font, sub_size)
    canvas.drawCentredString(page_w / 2, y_sub, subtitle)
    canvas.setFont(branch_font, branch_size)
    canvas.drawCentredString(page_w / 2, y_branch, branch_line)

    if print_dt_str:
        canvas.setFont(FONT_REGULAR, 7)
        canvas.drawRightString(page_w - 10 * mm, y_bank, f"Print: {print_dt_str}")
        canvas.drawRightString(page_w - 10 * mm, y_sub, "@Md. Noushad Ahmed")

    canvas.setFont(FONT_BOLD, 11)
    canvas.drawCentredString(page_w / 2, page_h - 39.5 * mm, title_text)
    canvas.setLineWidth(0.5)
    canvas.line(10 * mm, page_h - 41.5 * mm, page_w - 10 * mm, page_h - 41.5 * mm)
    canvas.setFont(FONT_REGULAR, 7)
    canvas.drawRightString(page_w - 10 * mm, 8 * mm, f"Page {doc.page}")
    canvas.restoreState()

def _fmt_cell(key, val):
    if val is None or str(val).strip() in ("", "None"):
        return ""
    if key in _AMOUNT_KEYS:
        try:
            return f"{float(str(val).replace(',', '')):,.0f}"
        except (TypeError, ValueError):
            return str(val)
    return str(val)


_SL_WEIGHT = 0.45


def _build_table(data_rows, page_w, cell_style, header_style, start_serial=1):
    header_row = [Paragraph("Sl.", header_style)] + [Paragraph(h, header_style) for _, h in COLUMNS]
    table_data = [header_row]
    for i, d in enumerate(data_rows):
        row = [Paragraph(str(start_serial + i), cell_style)] + \
              [Paragraph(_fmt_cell(k, d.get(k)), cell_style) for k, _ in COLUMNS]
        table_data.append(row)

    avail_width = page_w - 16 * mm
    total_weight = _SL_WEIGHT + sum(_COLUMN_WEIGHTS.get(k, 1.0) for k, _ in COLUMNS)
    col_widths = [avail_width * _SL_WEIGHT / total_weight] + \
                 [avail_width * _COLUMN_WEIGHTS.get(k, 1.0) / total_weight for k, _ in COLUMNS]
    t = Table(table_data, colWidths=col_widths, repeatRows=1)
    t.setStyle(TableStyle([
    ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#F2F2F2")),
    ("LINEBELOW", (0, 0), (-1, 0), 1, colors.black),   # হেডারের নিচে মোটা কালো লাইন
    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F2F2F2")]),  # খাঁটি ধূসর, কোনো রঙ না
    ("TOPPADDING", (0, 0), (-1, -1), 2),
    ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
]))
    return t


def generate_report_pdf(rows, out_path, branch_name, title_text,
                         bank_name=BANK_NAME_DEFAULT, logo_path=None, grouped=False, summary=None):
    """
    rows:
        grouped=False হলে -- filter_* ফাংশনের রেজাল্ট (flat list of dict)
        grouped=True হলে  -- group_by_union_village()-এর রেজাল্ট
    summary: {"label": "Balance"/"Due Amount", "count": int, "value": float} দিলে
        (শুধু grouped=False রিপোর্টে) রিপোর্টের একদম নিচে
        "Total Loans: N | Total <label>: X" লাইন বসবে।
        grouped=True হলে এই প্যারামিটার ব্যবহৃত হয় না -- প্রতি Union-এর সাবটোটাল
        (Loans/Balance/Due/Overdue/Rescheduled) এবং শেষে Grand Total,
        group_by_union_village()-এর subtotal dict থেকেই স্বয়ংক্রিয়ভাবে বসে।
    """
    page_w, page_h = landscape(A4)
    default_logo = os.path.join(os.path.dirname(__file__), "logo.png")
    if logo_path is None and os.path.exists(default_logo):
        logo_path = default_logo

    doc = SimpleDocTemplate(
        out_path, pagesize=landscape(A4),
        leftMargin=8 * mm, rightMargin=8 * mm,
        topMargin=45 * mm, bottomMargin=12 * mm,
    )

    styles = getSampleStyleSheet()
    cell_style = ParagraphStyle("cell", parent=styles["Normal"], fontSize=6.5, leading=8,
                                 fontName=FONT_REGULAR)
    header_style = ParagraphStyle(
        "hdr", parent=styles["Normal"], fontSize=7, leading=8,
        alignment=TA_CENTER, fontName=FONT_BOLD,
    )
    union_style = ParagraphStyle(
        "union", parent=styles["Heading4"], fontSize=11, spaceBefore=4, spaceAfter=3,
        fontName=FONT_BOLD,
    )
    subtotal_style = ParagraphStyle(
        "subtotal", parent=styles["Normal"], fontSize=9, fontName=FONT_BOLD, spaceAfter=8,
    )
    normal_style = ParagraphStyle("normal_txt", parent=styles["Normal"], fontName=FONT_REGULAR)

    elements = []
    if not grouped:
        if not rows:
            elements.append(Paragraph("No rows found for this filter.", normal_style))
        else:
            elements.append(_build_table(rows, page_w, cell_style, header_style))
            if summary is not None:
                elements.append(Spacer(1, 6))
                elements.append(Paragraph(
                    f"Total Loans: {summary['count']} &nbsp;&nbsp;|&nbsp;&nbsp; "
                    f"Total {summary['label']}: {summary['value']:,.0f}",
                    subtotal_style,
                ))
    else:
        if not rows:
            elements.append(Paragraph("No data found.", normal_style))
        grand = defaultdict(float)
        grand_keys = ["count", "balance", "due", "overdue_count", "overdue_balance",
                      "resch_count", "resch_balance"]
        for union, group_rows, subtotal in rows:
            elements.append(Paragraph(f"Union: {union}", union_style))
            elements.append(_build_table(group_rows, page_w, cell_style, header_style))
            elements.append(Paragraph(
                f"Total Loans: {subtotal.get('count', 0)} &nbsp;&nbsp;|&nbsp;&nbsp; "
                f"Total Balance: {subtotal.get('balance', 0):,.0f} &nbsp;&nbsp;|&nbsp;&nbsp; "
                f"Total Due: {subtotal.get('due', 0):,.0f} &nbsp;&nbsp;|&nbsp;&nbsp; "
                f"Total Overdue: {subtotal.get('overdue_count', 0)} "
                f"(Balance: {subtotal.get('overdue_balance', 0):,.0f}) &nbsp;&nbsp;|&nbsp;&nbsp; "
                f"Total Rescheduled: {subtotal.get('resch_count', 0)} "
                f"(Balance: {subtotal.get('resch_balance', 0):,.0f})",
                subtotal_style,
            ))
            elements.append(Spacer(1, 6))
            for k in grand_keys:
                grand[k] += subtotal.get(k, 0)
        if rows:
            elements.append(Paragraph(
                f"Grand Total Loans: {int(grand['count'])} &nbsp;&nbsp;|&nbsp;&nbsp; "
                f"Grand Total Balance: {grand['balance']:,.0f} &nbsp;&nbsp;|&nbsp;&nbsp; "
                f"Grand Total Due: {grand['due']:,.0f} &nbsp;&nbsp;|&nbsp;&nbsp; "
                f"Grand Total Overdue: {int(grand['overdue_count'])} "
                f"(Balance: {grand['overdue_balance']:,.0f}) &nbsp;&nbsp;|&nbsp;&nbsp; "
                f"Grand Total Rescheduled: {int(grand['resch_count'])} "
                f"(Balance: {grand['resch_balance']:,.0f})",
                subtotal_style,
            ))

    print_dt_str = datetime.now(BD_TZ).strftime("%d/%m/%Y %I:%M %p")
    header_fn = lambda c, d: _draw_header(c, d, bank_name, branch_name, logo_path, title_text, print_dt_str)
    doc.build(elements, onFirstPage=header_fn, onLaterPages=header_fn)
    return out_path
