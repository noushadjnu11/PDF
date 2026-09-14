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
    6. ৩-মাস ওভারভিউ রিপোর্ট (build_three_month_overview + generate_overview_report_pdf) --
       Regular → Overdue → Expired → Rescheduled → Due, এই ক্রমে ৫টা সেকশন, প্রতিটা
       Union-ভিত্তিক গ্রুপ করা, প্রতিটা Union সাব-টেবিল নতুন পেজে শুরু হয়। শুরুর তারিখ
       ডিফল্ট: চলতি মাসের ১-১০ হলে চলতি মাসের ১ তারিখ, নাহলে পরের মাসের ১ তারিখ
       (default_overdue_start_date) -- শেষের তারিখ ডিফল্ট শুরুর তারিখ + ৩ মাস
       (add_months), দুটোই UI-তে বদলানো যায়।

    প্রতিটা রিপোর্টে যেখানেই Union বাছাইয়ের অপশন আছে, সেখানে ঐচ্ছিকভাবে সেই
    Union(গুলো)-র ভেতরের Village-ও বাছাই করা যায় (খালি রাখলে সব Village আসবে)।

    প্রতিটা রিপোর্টের টেবিলে "Loan Case"-এর বামে একটা "Sl." (Serial) কলাম থাকে।
    সাধারণ রিপোর্টে এটা টানা ১, ২, ৩... — Union/Village রিপোর্টে প্রতিটা Union-এর
    জন্য আলাদাভাবে ১ থেকে শুরু হয়।

    প্রতিটা রিপোর্টেই বিদ্যমান (ডিফল্ট) সর্টিং-এর পাশাপাশি ঐচ্ছিকভাবে Loan Case
    অনুযায়ী সর্ট করা যায় (sort_by_loan_case=True) -- প্রিফিক্স-ভিত্তিক গ্রুপ করে
    (যেমন সব "OWN..." একসাথে) ছোট থেকে বড় সাজে।

    প্রতিটা রিপোর্ট PDF-এর পাশাপাশি Excel (.xlsx)-ও ডাউনলোড করা যায়
    (generate_report_excel) -- সেখানে ব্যাংকের নাম/লোগো/টাইটেল থাকে না, শুধু
    হেডার রো (Sl. + বাকি কলাম) আর তার নিচে ডেটা রো।

প্রতিটা PDF-এর প্রতি পেজেই উপরে ব্যাংকের লোগো + নাম + শাখা + রিপোর্ট-টাইটেল, এবং
ডান পাশে Print Date/Time + "@Md. Noushad Ahmed" থাকে।
"""
import os
import re
from collections import defaultdict
from datetime import date, datetime, timezone, timedelta

import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import (PageBreak, Paragraph, SimpleDocTemplate, Spacer,
                                 Table, TableStyle)

BANK_NAME_DEFAULT = "KARMASANGSTHAN BANK"
FONT_REGULAR = "Helvetica"
FONT_BOLD = "Helvetica-Bold"
BD_TZ = timezone(timedelta(hours=6))  # Bangladesh Standard Time, GMT+6 -- সার্ভারের timezone
                                       # যাই হোক না কেন, Print Date/Time সবসময় এই timezone-এ দেখানো হয়

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
    "prefixed_loan_case": 1.0,
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


def _s(v):
    """যেকোনো মান (None/সংখ্যা/তারিখ/string) নিরাপদে trimmed string-এ রূপান্তর করে।
    Excel-এর Union/Village/Loan Case-জাতীয় কলামে মাঝেমধ্যে সংখ্যা বা তারিখ টাইপ
    হিসেবে ডেটা এসে গেলে (d.get(...) or "").strip() AttributeError দিত -- এই
    হেল্পার সেটা এড়ায়।"""
    if v is None:
        return ""
    return str(v).strip()


_LOAN_CASE_RE = re.compile(r"^(\D*)(\d*)")


def _loan_case_sort_key(row):
    """Loan Case-কে অক্ষর-প্রিফিক্স + সংখ্যা অংশে ভেঙে সাজানোর key বানায়, যেমন
    'OWN2' < 'OWN10' (সংখ্যা হিসেবে, string হিসেবে না) -- একই প্রিফিক্সের (যেমন 'OWN')
    সব Loan Case একসাথে গ্রুপ হয়ে ছোট থেকে বড় সাজে, তারপর পরের প্রিফিক্স।"""
    s = _s(row.get("prefixed_loan_case"))
    m = _LOAN_CASE_RE.match(s)
    prefix = (m.group(1) or "").strip().upper() if m else ""
    num_str = (m.group(2) or "") if m else ""
    num = int(num_str) if num_str.isdigit() else 0
    return (prefix, num, s.upper())


def default_overdue_start_date(today=None):
    """Overdue-জাতীয় রিপোর্টের শুরুর তারিখ অটো-বসানোর নিয়ম: আজকের তারিখ চলতি মাসের
    ১-১০ এর মধ্যে হলে চলতি মাসের ১ তারিখ, নাহলে (১১ বা তার পরে হলে) পরের মাসের
    ১ তারিখ। ইউজার চাইলে UI-তে এটা বদলে দিতে পারবে (শুধু ডিফল্ট মান)।"""
    today = today or date.today()
    if today.day <= 10:
        return date(today.year, today.month, 1)
    if today.month == 12:
        return date(today.year + 1, 1, 1)
    return date(today.year, today.month + 1, 1)


def add_months(d, months):
    """d তারিখের সাথে months সংখ্যক মাস যোগ করে -- মাসের শেষ দিনের হিসাব ঠিক রেখে
    (যেমন 31 জানুয়ারি + 1 মাস = 28/29 ফেব্রুয়ারি)।"""
    month_index = d.month - 1 + months
    year = d.year + month_index // 12
    month = month_index % 12 + 1
    is_leap = year % 4 == 0 and (year % 100 != 0 or year % 400 == 0)
    days_in_month = [31, 29 if is_leap else 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    day = min(d.day, days_in_month[month - 1])
    return date(year, month, day)



def _apply_union_filter(rows, unions):
    """unions দেওয়া থাকলে শুধু সেই Union(গুলো)-র রো রাখে; না থাকলে সব রো ফেরত দেয়।"""
    union_set = {u.strip().lower() for u in (unions or []) if u and u.strip()}
    if not union_set:
        return rows
    return [d for d in rows if _s(d.get("union")).lower() in union_set]


def _apply_village_filter(rows, villages):
    """villages দেওয়া থাকলে শুধু সেই Village(গুলো)-র রো রাখে; না থাকলে সব রো ফেরত দেয়।"""
    village_set = {v.strip().lower() for v in (villages or []) if v and v.strip()}
    if not village_set:
        return rows
    return [d for d in rows if _s(d.get("village")).lower() in village_set]


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
        u_key = _s(d.get("union")).lower()
        if u_key not in norm_map:
            continue
        v_set = norm_map[u_key]
        if not v_set or _s(d.get("village")).lower() in v_set:
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
    return sorted({_s(d.get("union")) for d in rows if d.get("union")})


def get_villages_for_unions(rows, unions=None):
    """দেওয়া Union(গুলো)-র মধ্যে থাকা ইউনিক, sorted Village-এর তালিকা।
    unions খালি থাকলে সব রো-এর Village ফেরত দেয়।"""
    base = _apply_union_filter(rows, unions)
    return sorted({_s(d.get("village")) for d in base if d.get("village")})


def filter_overdue(rows, start_date, end_date, union_village_map=None, sort_by_loan_case=False):
    """start_date <= overdue_date <= end_date (দুই পাশই inclusive) -- ঐচ্ছিক
    per-Union Village filter সহ -- ডিফল্ট sort: oldest → newest (overdue date অনুযায়ী)।
    sort_by_loan_case=True দিলে তার বদলে Loan Case অনুযায়ী (প্রিফিক্স-ভিত্তিক গ্রুপ করে
    ছোট থেকে বড়) সাজানো হয়।"""
    out = [
        d for d in rows
        if (dt := parse_ddmmyyyy(d.get("overdue_date"))) and start_date <= dt <= end_date
    ]
    out = _apply_union_village_map_filter(out, union_village_map)
    if sort_by_loan_case:
        out.sort(key=_loan_case_sort_key)
    else:
        out.sort(key=lambda d: parse_ddmmyyyy(d.get("overdue_date")))
    return out


def filter_expired(rows, before, union_village_map=None, sort_by_loan_case=False):
    """overdue_date <= before (তারিখসহ, inclusive) -- ঐচ্ছিক per-Union Village filter সহ --
    ডিফল্ট sort: Union অনুযায়ী, প্রতি Union-এর ভেতরে oldest → newest।
    sort_by_loan_case=True দিলে তার বদলে Loan Case অনুযায়ী সাজানো হয়।"""
    out = [d for d in rows if (dt := parse_ddmmyyyy(d.get("overdue_date"))) and dt <= before]
    out = _apply_union_village_map_filter(out, union_village_map)
    if sort_by_loan_case:
        out.sort(key=_loan_case_sort_key)
    else:
        out.sort(key=lambda d: (_s(d.get("union")).lower(), parse_ddmmyyyy(d.get("overdue_date"))))
    return out


def filter_rescheduled(rows, after, union_village_map=None, sort_by_loan_case=False):
    """overdue_date > after এবং Reschedule No. > 0 -- ঐচ্ছিক per-Union Village filter সহ --
    ডিফল্ট sort: Union অনুযায়ী, প্রতি Union-এর ভেতরে oldest → newest।
    sort_by_loan_case=True দিলে তার বদলে Loan Case অনুযায়ী সাজানো হয়।"""
    out = [
        d for d in rows
        if (dt := parse_ddmmyyyy(d.get("overdue_date"))) and dt > after and _num(d.get("reschedule_no")) > 0
    ]
    out = _apply_union_village_map_filter(out, union_village_map)
    if sort_by_loan_case:
        out.sort(key=_loan_case_sort_key)
    else:
        out.sort(key=lambda d: (_s(d.get("union")).lower(), parse_ddmmyyyy(d.get("overdue_date"))))
    return out


def filter_due_amount(rows, union_village_map=None, sort_by_loan_case=False):
    """যেসব রো-তে Due Amount উপলব্ধ (> 0), শুধু সেগুলো -- ঐচ্ছিক per-Union Village filter সহ --
    ডিফল্ট sort: Union অনুযায়ী, প্রতি Union-এর ভেতরে oldest → newest।
    sort_by_loan_case=True দিলে তার বদলে Loan Case অনুযায়ী সাজানো হয়।"""
    out = [d for d in rows if _num(d.get("due_amount")) > 0]
    out = _apply_union_village_map_filter(out, union_village_map)
    if sort_by_loan_case:
        out.sort(key=_loan_case_sort_key)
    else:
        out.sort(key=lambda d: (_s(d.get("union")).lower(),
                                 parse_ddmmyyyy(d.get("overdue_date")) or date.min))
    return out


def group_by_union_village(rows, union_village_map=None, ref_date=None, sort_by_loan_case=False):
    """Union অনুযায়ী গ্রুপ (Union নাম অনুযায়ী sorted), প্রতি গ্রুপের ভেতরে ডিফল্ট sort:
    Village অনুযায়ী। sort_by_loan_case=True দিলে প্রতি Union-গ্রুপের ভেতরে তার বদলে
    Loan Case অনুযায়ী (প্রিফিক্স-ভিত্তিক গ্রুপ করে ছোট থেকে বড়) সাজানো হয়।
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
        union = _s(d.get("union")) or "Unknown"
        groups[union].append(d)

    result = []
    for union in sorted(groups.keys()):
        if sort_by_loan_case:
            group_rows = sorted(groups[union], key=_loan_case_sort_key)
        else:
            group_rows = sorted(groups[union], key=lambda d: _s(d.get("village")))

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


OVERVIEW_SECTIONS = [
    ("regular", "Regular Loan"),
    ("overdue", "Overdue Loan"),
    ("expired", "Expired Loan"),
    ("rescheduled", "Rescheduled Loan"),
    ("due", "Due Loan"),
]


def build_three_month_overview(rows, union_village_map=None, start_date=None, end_date=None,
                                sort_by_loan_case=None):
    """"৩-মাস ওভারভিউ" রিপোর্টের জন্য ৫টা ক্যাটাগরি -- Regular, Overdue, Expired,
    Rescheduled, Due -- Union-ভিত্তিক গ্রুপ করে রিটার্ন করে।
        Regular Loan     -- overdue_date > start_date (অর্থাৎ Expired বাদে বাকি সব)
        Overdue Loan     -- start_date <= overdue_date <= end_date
        Expired Loan     -- overdue_date <= start_date (তারিখসহ)
        Rescheduled Loan -- overdue_date > start_date এবং Reschedule No. > 0
        Due Loan         -- Due Amount > 0
    sort_by_loan_case: প্রতিটা ক্যাটাগরির জন্য আলাদা sort পছন্দ, dict যেমন
        {"regular": False, "overdue": True, ...} (না দেওয়া key/dict মানে ডিফল্ট
        False -- সেক্ষেত্রে প্রতি Union-এর ভেতরে Overdue Date অনুযায়ী সাজে)।
    রিটার্ন: {"regular": [...], "overdue": [...], "expired": [...],
              "rescheduled": [...], "due": [...]}
    প্রতিটা মান group_by_union_village()-এর মতো [(union, rows, {"count","balance"}), ...]।
    """
    sort_by_loan_case = sort_by_loan_case or {}
    if start_date is None:
        start_date = default_overdue_start_date()
    if end_date is None:
        end_date = add_months(start_date, 3)

    filtered = _apply_union_village_map_filter(rows, union_village_map)

    def _group(flat_rows, sort_lc):
        groups = defaultdict(list)
        for d in flat_rows:
            union = _s(d.get("union")) or "Unknown"
            groups[union].append(d)
        result = []
        for union in sorted(groups.keys()):
            if sort_lc:
                grows = sorted(groups[union], key=_loan_case_sort_key)
            else:
                grows = sorted(groups[union], key=lambda d: parse_ddmmyyyy(d.get("overdue_date")) or date.min)
            balance = sum(_num(d.get("bal_total")) for d in grows)
            result.append((union, grows, {"count": len(grows), "balance": balance}))
        return result

    regular_rows = [
        d for d in filtered if (dt := parse_ddmmyyyy(d.get("overdue_date"))) and dt > start_date
    ]
    overdue_rows = [
        d for d in filtered
        if (dt := parse_ddmmyyyy(d.get("overdue_date"))) and start_date <= dt <= end_date
    ]
    expired_rows = [
        d for d in filtered if (dt := parse_ddmmyyyy(d.get("overdue_date"))) and dt <= start_date
    ]
    resch_rows = [
        d for d in filtered
        if (dt := parse_ddmmyyyy(d.get("overdue_date"))) and dt > start_date
        and _num(d.get("reschedule_no")) > 0
    ]
    due_rows = [d for d in filtered if _num(d.get("due_amount")) > 0]

    return {
        "regular": _group(regular_rows, sort_by_loan_case.get("regular", False)),
        "overdue": _group(overdue_rows, sort_by_loan_case.get("overdue", False)),
        "expired": _group(expired_rows, sort_by_loan_case.get("expired", False)),
        "rescheduled": _group(resch_rows, sort_by_loan_case.get("rescheduled", False)),
        "due": _group(due_rows, sort_by_loan_case.get("due", False)),
    }

def build_output_filename(report_key, union_village_map=None, start=None, end=None,
                           single_date=None, ext="pdf"):
    """ফাইলের নাম dynamic ভাবে বানায়, যেমন:
    Overdue_Joynagar-Partial_Kashipur_15-03-2025_to_30-06-2026.pdf
    Expired_AllUnion_29-08-2026.xlsx
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
    return "_".join(parts) + f".{ext}"


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


def _excel_cell(key, val):
    """Excel-এর সেলে বসানোর জন্য মান রূপান্তর করে -- Amount কলাম হলে সংখ্যা (float/int)
    রিটার্ন করে (যাতে Excel-এ যোগ/বিয়োগ করা যায়, কমা-বসানো string না), বাকি সব
    কলাম আসল string/None-ই থাকে।"""
    if val is None or str(val).strip() in ("", "None"):
        return None
    if key in _AMOUNT_KEYS:
        try:
            n = float(str(val).replace(",", ""))
            return int(n) if n == int(n) else n
        except (TypeError, ValueError):
            return val
    return val


def generate_report_excel(rows, out_path, grouped=False):
    """ব্যাংকের নাম/লোগো/শাখা/টাইটেল ছাড়া শুধু হেডার রো + তার নিচে ডেটা রো সহ একটা
    .xlsx ফাইল বানায় -- PDF টেবিলের মতোই একই কলাম-অর্ডারে, বামে Sl. কলামসহ।
    rows:
        grouped=False হলে -- filter_* ফাংশনের রেজাল্ট (flat list of dict)
        grouped=True হলে  -- group_by_union_village()-এর রেজাল্ট -- এক্ষেত্রে সব
        Union-এর রো একটাই টানা টেবিলে (Union কলাম দিয়ে চেনা যাবে) বসে, Sl. পুরো
        শিটজুড়ে টানা ১, ২, ৩...; কোনো সাবটোটাল/গ্র্যান্ড-টোটাল রো থাকে না।"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Report"

    header = ["Sl."] + [h for _, h in COLUMNS]
    ws.append(header)
    for cell in ws[1]:
        cell.font = Font(bold=True)

    if grouped:
        flat_rows = []
        for _union, group_rows, _subtotal in rows:
            flat_rows.extend(group_rows)
    else:
        flat_rows = rows

    for i, d in enumerate(flat_rows, start=1):
        ws.append([i] + [_excel_cell(k, d.get(k)) for k, _ in COLUMNS])

    widths = [6] + [max(10, len(h) + 2) for _, h in COLUMNS]
    for idx, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = w

    ws.freeze_panes = "A2"
    wb.save(out_path)
    return out_path


def generate_overview_report_pdf(overview_data, out_path, branch_name, title_text,
                                  bank_name=BANK_NAME_DEFAULT, logo_path=None):
    """overview_data: build_three_month_overview()-এর রেজাল্ট। Regular → Overdue →
    Expired → Rescheduled → Due -- এই ক্রমে ৫টা সেকশন, প্রতিটার ভেতরে Union-ভিত্তিক
    সাব-টেবিল। প্রতিটা Union সাব-টেবিল (প্রতিটা সেকশনেরও প্রথমটা সহ) নতুন পেজে শুরু
    হয়, যাতে কোনো একটা পেজে দুইটা টেবিল-হেডার-রো (Sl., Loan Case, ...) একসাথে না
    থাকে -- প্রতিটা পেজে শুধু একটাই টাইটেল রো।"""
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
    section_style = ParagraphStyle(
        "section", parent=styles["Heading3"], fontSize=13, spaceBefore=0, spaceAfter=6,
        fontName=FONT_BOLD,
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
    first_section = True
    for key, label in OVERVIEW_SECTIONS:
        groups = overview_data.get(key) or []
        if not first_section:
            elements.append(PageBreak())
        first_section = False

        if not groups:
            elements.append(Paragraph(label, section_style))
            elements.append(Paragraph("No data found.", normal_style))
            continue

        grand_count, grand_balance = 0, 0.0
        first_union = True
        for union, group_rows, subtotal in groups:
            if not first_union:
                elements.append(PageBreak())
            first_union = False
            elements.append(Paragraph(label, section_style))
            elements.append(Paragraph(f"Union: {union}", union_style))
            elements.append(_build_table(group_rows, page_w, cell_style, header_style))
            elements.append(Paragraph(
                f"Total Loans: {subtotal.get('count', 0)} &nbsp;&nbsp;|&nbsp;&nbsp; "
                f"Total Balance: {subtotal.get('balance', 0):,.0f}",
                subtotal_style,
            ))
            elements.append(Spacer(1, 6))
            grand_count += subtotal.get("count", 0)
            grand_balance += subtotal.get("balance", 0)

        elements.append(Paragraph(
            f"{label} — Grand Total Loans: {grand_count} &nbsp;&nbsp;|&nbsp;&nbsp; "
            f"Grand Total Balance: {grand_balance:,.0f}",
            subtotal_style,
        ))

    print_dt_str = datetime.now(BD_TZ).strftime("%d/%m/%Y %I:%M %p")
    header_fn = lambda c, d: _draw_header(c, d, bank_name, branch_name, logo_path, title_text, print_dt_str)
    doc.build(elements, onFirstPage=header_fn, onLaterPages=header_fn)
    return out_path


def generate_overview_excel(overview_data, out_path):
    """overview_data: build_three_month_overview()-এর রেজাল্ট। এক শিটে ৫টা সেকশন
    (Regular → Overdue → Expired → Rescheduled → Due), প্রতিটার আগে একটা বোল্ড লেবেল
    রো, তারপর হেডার রো (Sl. + বাকি কলাম + Union), তারপর ডেটা রো -- Union অনুযায়ী
    গ্রুপ করা, প্রতি Union-এর জন্য Sl. নতুন করে ১ থেকে শুরু। ব্যাংকের নাম/লোগো/টাইটেল
    নেই।"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "3-Month Overview"

    r = 1
    for key, label in OVERVIEW_SECTIONS:
        groups = overview_data.get(key) or []
        label_cell = ws.cell(row=r, column=1, value=label)
        label_cell.font = Font(bold=True, size=13)
        r += 1

        header = ["Sl."] + [h for _, h in COLUMNS] + ["Union"]
        for idx, h in enumerate(header, start=1):
            c = ws.cell(row=r, column=idx, value=h)
            c.font = Font(bold=True)
        r += 1

        for union, group_rows, _subtotal in groups:
            for i, d in enumerate(group_rows, start=1):
                row_vals = [i] + [_excel_cell(k, d.get(k)) for k, _ in COLUMNS] + [union]
                for idx, v in enumerate(row_vals, start=1):
                    ws.cell(row=r, column=idx, value=v)
                r += 1
        r += 2  # পরের সেকশনের আগে ফাঁকা গ্যাপ

    widths = [6] + [max(10, len(h) + 2) for _, h in COLUMNS] + [14]
    for idx, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = w

    wb.save(out_path)
    return out_path
