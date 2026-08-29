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
    ("overdue_date", "Overdue Date"),
    ("installment", "Installment"),
    ("bal_principal", "Principal"),
    ("bal_interest", "Interest"),
    ("bal_total", "Balance Total"),
    ("due_amount", "Due Amount"),
    ("reschedule_no", "Reschedule No."),
    ("blank_col", "Comment"),
]

_AMOUNT_KEYS = {"installment", "bal_principal", "bal_interest", "bal_total", "due_amount"}


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


def filter_expired(rows, before):
    """overdue_date < before -- oldest → newest sort।"""
    out = [d for d in rows if (dt := parse_ddmmyyyy(d.get("overdue_date"))) and dt < before]
    out.sort(key=lambda d: parse_ddmmyyyy(d.get("overdue_date")))
    return out


def filter_rescheduled(rows, after):
    """overdue_date > after এবং Reschedule No. > 0 -- oldest → newest sort।"""
    out = [
        d for d in rows
        if (dt := parse_ddmmyyyy(d.get("overdue_date"))) and dt > after and _num(d.get("reschedule_no")) > 0
    ]
    out.sort(key=lambda d: parse_ddmmyyyy(d.get("overdue_date")))
    return out


def group_by_union_village(rows):
    """Union অনুযায়ী গ্রুপ (Union নাম অনুযায়ী sorted), প্রতি গ্রুপে Village অনুযায়ী sort,
    এবং প্রতি গ্রুপের সাথে (loan_count, balance_total) সাবটোটাল।
    রিটার্ন: [(union_name, [row, ...], {"count":.., "balance":..}), ...]"""
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


# =============================================================================
# ৩. PDF জেনারেশন (A4 Landscape, প্রতি পেজে ব্যাংক-হেডার)
# =============================================================================
def _draw_header(canvas, doc, bank_name, branch_name, logo_path, title_text):
    canvas.saveState()
    page_w, page_h = landscape(A4)

    subtitle = "A State Owned Financial Institution"
    branch_line = f"{branch_name}"

    # প্রতিটা লাইনের ফন্ট+সাইজ যা বসবে, সেই অনুযায়ী প্রস্থ মাপা
    bank_font, bank_size = FONT_BOLD, 13
    sub_font, sub_size = FONT_REGULAR, 8
    branch_font, branch_size = FONT_REGULAR, 9

    widths = [
        canvas.stringWidth(bank_name, bank_font, bank_size),
        canvas.stringWidth(subtitle, sub_font, sub_size),
        canvas.stringWidth(branch_line, branch_font, branch_size),
    ]
    max_text_width = max(widths)
    text_left_edge = page_w / 2 - max_text_width / 2   # সবচেয়ে চওড়া লাইনের বাম-প্রান্ত

    # লোগো এই বাম-প্রান্তের ঠিক গা ঘেঁষে, বাম দিকে
    logo_size = 16 * mm
    gap = 3 * mm
    logo_x = text_left_edge - gap - logo_size

    if logo_path and os.path.exists(logo_path):
        canvas.drawImage(
            logo_path, logo_x, page_h - 26 * mm,
            width=logo_size, height=logo_size, mask="auto", preserveAspectRatio=True,
        )

    # টেক্সট আগের মতোই পেজ-সেন্টার বরাবর
    canvas.setFont(bank_font, bank_size)
    canvas.drawCentredString(page_w / 2, page_h - 20 * mm, bank_name)

    canvas.setFont(sub_font, sub_size)
    canvas.drawCentredString(page_w / 2, page_h - 25 * mm, subtitle)

    canvas.setFont(branch_font, branch_size)
    canvas.drawCentredString(page_w / 2, page_h - 30 * mm, branch_line)

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


def _build_table(data_rows, page_w, cell_style, header_style):
    table_data = [[Paragraph(h, header_style) for _, h in COLUMNS]]
    for d in data_rows:
        table_data.append([Paragraph(_fmt_cell(k, d.get(k)), cell_style) for k, _ in COLUMNS])

    avail_width = page_w - 16 * mm
    col_widths = [avail_width / len(COLUMNS)] * len(COLUMNS)
    t = Table(table_data, colWidths=col_widths, repeatRows=1)
    t.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#D9E1F2")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F5F7FB")]),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    return t


def generate_report_pdf(rows, out_path, branch_name, title_text,
                         bank_name=BANK_NAME_DEFAULT, logo_path=None, grouped=False):
    """
    rows:
        grouped=False হলে -- filter_* ফাংশনের রেজাল্ট (flat list of dict)
        grouped=True হলে  -- group_by_union_village()-এর রেজাল্ট
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

    header_fn = lambda c, d: _draw_header(c, d, bank_name, branch_name, logo_path, title_text)
    doc.build(elements, onFirstPage=header_fn, onLaterPages=header_fn)
    return out_path
