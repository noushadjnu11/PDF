# -*- coding: utf-8 -*-
"""
report_generator.py
============================================================
merged_loan_report.xlsx থেকে বিভিন্ন ধরনের A4-Landscape ব্যাংক PDF রিপোর্ট বানানোর মডিউল।

সাপোর্টেড রিপোর্ট:
    1. Overdue Loan up to <date>         -- overdue_date <= date, oldest→newest
    2. (একই, কিন্তু) নির্দিষ্ট এক/একাধিক Union filter সহ
    3. Expired Loan List up to <date>    -- overdue_date < date, oldest→newest
    4. Rescheduled Loan up to <date>     -- overdue_date > date এবং Reschedule No. > 0
    5. Union/Village-ভিত্তিক গ্রুপড রিপোর্ট (সংশোধিত Excel থেকে), প্রতি Union-এর
       শেষে মোট লোন সংখ্যা ও মোট ব্যালেন্স সাবটোটাল সহ।

প্রতিটা PDF-এর প্রতি পেজেই উপরে ব্যাংকের লোগো + নাম + শাখা + রিপোর্ট-টাইটেল থাকে।
"""
import os
import re
from collections import defaultdict
from datetime import date, datetime

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

# Excel-এর কলাম অর্ডারের সাথে হুবহু মিল রেখে (pdf_processor.write_excel দেখুন)
COLUMNS = [
    ("prefixed_loan_case", "Loan Case"),
    ("borrower", "Borrower"),
    ("father", "Father"),
    ("spouse", "Spouse"),
    ("village", "Village"),
    ("union", "Union"),
    ("phone", "Phone"),
    ("overdue_date", "Overdue"),
    ("installment", "Instl."),
    ("bal_principal", "Princ."),
    ("bal_interest", "Interest"),
    ("bal_total", "Balance"),
    ("due_amount", "Due"),
    ("reschedule_no", "Resch."),
    ("blank_col", "Comment"),
]

_AMOUNT_KEYS = {"installment", "bal_principal", "bal_interest", "bal_total", "due_amount"}

# কলামগুলোর আপেক্ষিক প্রস্থ -- সংখ্যা/তারিখ কলাম সরু, নাম/ঠিকানা কলাম চওড়া
_COLUMN_WEIGHTS = {
    "prefixed_loan_case": 1.0,
    "borrower": 1.3,
    "father": 1.3,
    "spouse": 1.3,
    "village": 1.0,
    "union": 1.0,
    "phone": 1.15,
    "overdue_date": 1.0,
    "installment": 0.85,
    "bal_principal": 0.8,
    "bal_interest": 0.7,
    "bal_total": 0.8,
    "due_amount": 0.65,
    "reschedule_no": 0.3,
    "blank_col": 0.9,
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
def filter_overdue(rows, start_date, end_date):
    """start_date <= overdue_date <= end_date (দুই পাশই inclusive) -- oldest → newest sort।"""
    out = [
        d for d in rows
        if (dt := parse_ddmmyyyy(d.get("overdue_date"))) and start_date <= dt <= end_date
    ]
    out.sort(key=lambda d: parse_ddmmyyyy(d.get("overdue_date")))
    return out


def filter_union_overdue(rows, start_date, end_date, unions):
    """filter_overdue()-এর ফলাফল থেকে শুধু দেওয়া Union(গুলো)-র রো।"""
    union_set = {u.strip().lower() for u in unions if u and u.strip()}
    base = filter_overdue(rows, start_date, end_date)
    if not union_set:
        return base
    return [d for d in base if (d.get("union") or "").strip().lower() in union_set]


def _apply_union_filter(rows, unions):
    """unions দেওয়া থাকলে শুধু সেই Union(গুলো)-র রো রাখে; না থাকলে সব রো ফেরত দেয়।"""
    union_set = {u.strip().lower() for u in (unions or []) if u and u.strip()}
    if not union_set:
        return rows
    return [d for d in rows if (d.get("union") or "").strip().lower() in union_set]


def filter_expired(rows, before, unions=None):
    """overdue_date < before -- ঐচ্ছিক Union filter সহ -- Union অনুযায়ী সাজানো,
    প্রতি Union-এর ভেতরে oldest → newest।"""
    out = [d for d in rows if (dt := parse_ddmmyyyy(d.get("overdue_date"))) and dt < before]
    out = _apply_union_filter(out, unions)
    out.sort(key=lambda d: ((d.get("union") or "").strip().lower(), parse_ddmmyyyy(d.get("overdue_date"))))
    return out


def filter_rescheduled(rows, after, unions=None):
    """overdue_date > after এবং Reschedule No. > 0 -- ঐচ্ছিক Union filter সহ --
    Union অনুযায়ী সাজানো, প্রতি Union-এর ভেতরে oldest → newest।"""
    out = [
        d for d in rows
        if (dt := parse_ddmmyyyy(d.get("overdue_date"))) and dt > after and _num(d.get("reschedule_no")) > 0
    ]
    out = _apply_union_filter(out, unions)
    out.sort(key=lambda d: ((d.get("union") or "").strip().lower(), parse_ddmmyyyy(d.get("overdue_date"))))
    return out


def filter_due_amount(rows, unions=None):
    """যেসব রো-তে Due Amount উপলব্ধ (> 0), শুধু সেগুলো -- ঐচ্ছিক Union filter সহ --
    Union অনুযায়ী সাজানো, প্রতি Union-এর ভেতরে oldest → newest।"""
    out = [d for d in rows if _num(d.get("due_amount")) > 0]
    out = _apply_union_filter(out, unions)
    out.sort(key=lambda d: ((d.get("union") or "").strip().lower(),
                             parse_ddmmyyyy(d.get("overdue_date")) or date.min))
    return out


def group_by_union_village(rows, unions=None):
    """Union অনুযায়ী গ্রুপ (Union নাম অনুযায়ী sorted), প্রতি গ্রুপে Village অনুযায়ী sort,
    এবং প্রতি গ্রুপের সাথে (loan_count, balance_total) সাবটোটাল। ঐচ্ছিক Union filter সহ।
    রিটার্ন: [(union_name, [row, ...], {"count":.., "balance":..}), ...]"""
    rows = _apply_union_filter(rows, unions)
    groups = defaultdict(list)
    for d in rows:
        union = (d.get("union") or "Unknown").strip() or "Unknown"
        groups[union].append(d)

    result = []
    for union in sorted(groups.keys()):
        group_rows = sorted(groups[union], key=lambda d: (d.get("village") or "").strip())
        subtotal = {
            "count": len(group_rows),
            "balance": sum(_num(d.get("bal_total")) for d in group_rows),
        }
        result.append((union, group_rows, subtotal))
    return result


def build_output_filename(report_key, unions=None, start=None, end=None, single_date=None):
    """PDF ফাইলের নাম dynamic ভাবে বানায়, যেমন:
    Overdue_Joynagar_15-03-2025_to_30-06-2026.pdf
    Expired_AllUnion_29-08-2026.pdf
    """
    def _clean(s):
        return re.sub(r"[^A-Za-z0-9]+", "", s) or "Union"

    parts = [report_key]
    if unions:
        parts.append("-".join(_clean(u) for u in unions))
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
def _draw_header(canvas, doc, bank_name, branch_name, logo_path, title_text):
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


def _build_table(data_rows, page_w, cell_style, header_style):
    table_data = [[Paragraph(h, header_style) for _, h in COLUMNS]]
    for d in data_rows:
        table_data.append([Paragraph(_fmt_cell(k, d.get(k)), cell_style) for k, _ in COLUMNS])

    avail_width = page_w - 16 * mm
    total_weight = sum(_COLUMN_WEIGHTS.get(k, 1.0) for k, _ in COLUMNS)
    col_widths = [avail_width * _COLUMN_WEIGHTS.get(k, 1.0) / total_weight for k, _ in COLUMNS]
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
        রিপোর্টের একদম নিচে "Total Loans: N | Total <label>: X" লাইন বসবে
        (grouped=True হলে প্রতি Union-এর সাবটোটালের নিচে, সব গ্রুপ শেষে গ্র্যান্ড-টোটাল হিসেবে)।
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
        for union, group_rows, subtotal in rows:
            elements.append(Paragraph(f"Union: {union}", union_style))
            elements.append(_build_table(group_rows, page_w, cell_style, header_style))
            elements.append(Paragraph(
                f"Total Loans: {subtotal['count']} &nbsp;&nbsp;|&nbsp;&nbsp; "
                f"Total Balance: {subtotal['balance']:,.0f}",
                subtotal_style,
            ))
            elements.append(Spacer(1, 6))
        if summary is not None and rows:
            elements.append(Paragraph(
                f"Grand Total Loans: {summary['count']} &nbsp;&nbsp;|&nbsp;&nbsp; "
                f"Grand Total {summary['label']}: {summary['value']:,.0f}",
                subtotal_style,
            ))

    header_fn = lambda c, d: _draw_header(c, d, bank_name, branch_name, logo_path, title_text)
    doc.build(elements, onFirstPage=header_fn, onLaterPages=header_fn)
    return out_path
