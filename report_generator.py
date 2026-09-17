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
    ("overdue", "Overdue Loan"),
    ("expired", "Expired Loan"),
    ("rescheduled", "Rescheduled Loan"),
    ("due", "Due Loan"),
]


def build_overview_report_data(rows, union_village_map=None, cutoff_date=None,
                                overdue_start=None, overdue_end=None, sort_by_loan_case=None):
    """"Overview Report"-এর জন্য ডেটা তৈরি করে।
        Overall Loan Information -- Union-ভিত্তিক সারাংশ (কোনো লোন-লিস্টিং না):
            প্রতি Union-এ Loan Count, Balance, Expired Count+Balance,
            Rescheduled Count+Balance, Due Count+Balance -- শেষে Grand Total।
        Overdue Loan      -- overdue_start <= overdue_date <= overdue_end
        Expired Loan      -- overdue_date <= cutoff_date (তারিখসহ, inclusive)
        Rescheduled Loan  -- overdue_date > cutoff_date (exclusive) এবং Reschedule No. > 0
        Due Loan          -- Due Amount > 0
    এই ৪টা টেবিল Union-ভিত্তিক গ্রুপ না করে ফ্ল্যাট লিস্ট হিসেবে থাকে -- ডিফল্ট সর্ট
    Union তারপর Village অনুযায়ী; sort_by_loan_case-এ True দিলে সেই ক্যাটাগরির জন্য
    Loan Case অনুযায়ী সাজে।

    cutoff_date: Expired/Rescheduled দুটোর জন্যই এই একটা তারিখ ব্যবহার হয়। না দিলে
        default_overdue_start_date()।
    overdue_start/overdue_end: শুধু Overdue টেবিলের রেঞ্জ (cutoff_date থেকে সম্পূর্ণ
        আলাদা, ইউজার এডিট করতে পারবে)। না দিলে overdue_start=cutoff_date,
        overdue_end=add_months(cutoff_date, 3)।
    sort_by_loan_case: dict, যেমন {"overdue": True, "expired": False, ...}।

    রিটার্ন:
        {
          "cutoff_date", "overdue_start", "overdue_end",
          "overall_summary": [(union, {count,balance,expired_count,expired_balance,
                                        resch_count,resch_balance,due_count,due_balance}), ...],
          "overall_grand": {...একই key-গুলোর যোগফল...},
          "overdue": [row, ...], "expired": [...], "rescheduled": [...], "due": [...],
        }
    """
    sort_by_loan_case = sort_by_loan_case or {}
    if cutoff_date is None:
        cutoff_date = default_overdue_start_date()
    if overdue_start is None:
        overdue_start = cutoff_date
    if overdue_end is None:
        overdue_end = add_months(overdue_start, 3)

    filtered = _apply_union_village_map_filter(rows, union_village_map)

    def _sorted_flat(flat_rows, key):
        if sort_by_loan_case.get(key, False):
            return sorted(flat_rows, key=_loan_case_sort_key)
        return sorted(flat_rows, key=lambda d: (_s(d.get("union")).lower(), _s(d.get("village")).lower()))

    overdue_rows = [
        d for d in filtered
        if (dt := parse_ddmmyyyy(d.get("overdue_date"))) and overdue_start <= dt <= overdue_end
    ]
    expired_rows = [
        d for d in filtered if (dt := parse_ddmmyyyy(d.get("overdue_date"))) and dt <= cutoff_date
    ]
    resch_rows = [
        d for d in filtered
        if (dt := parse_ddmmyyyy(d.get("overdue_date"))) and dt > cutoff_date
        and _num(d.get("reschedule_no")) > 0
    ]
    due_rows = [d for d in filtered if _num(d.get("due_amount")) > 0]

    # --- Overall Loan Information: Union-ভিত্তিক সারাংশ ---
    def _by_union(flat_rows):
        out = defaultdict(list)
        for d in flat_rows:
            out[_s(d.get("union")) or "Unknown"].append(d)
        return out

    all_by_union = _by_union(filtered)
    expired_by_union = _by_union(expired_rows)
    resch_by_union = _by_union(resch_rows)
    due_by_union = _by_union(due_rows)

    overall_summary = []
    grand = defaultdict(float)
    for union in sorted(all_by_union.keys()):
        u_rows = all_by_union[union]
        u_expired = expired_by_union.get(union, [])
        u_resch = resch_by_union.get(union, [])
        u_due = due_by_union.get(union, [])
        row = {
            "count": len(u_rows),
            "balance": sum(_num(d.get("bal_total")) for d in u_rows),
            "expired_count": len(u_expired),
            "expired_balance": sum(_num(d.get("bal_total")) for d in u_expired),
            "resch_count": len(u_resch),
            "resch_balance": sum(_num(d.get("bal_total")) for d in u_resch),
            "due_count": len(u_due),
            "due_balance": sum(_num(d.get("bal_total")) for d in u_due),
        }
        overall_summary.append((union, row))
        for k, v in row.items():
            grand[k] += v

    # --- Village/Union Wise: সিলেক্টেড স্কোপের সব লোন (কোনো ক্যাটাগরি-ফিল্টার ছাড়া)
    # একসাথে মিলিয়ে Union → Village অনুযায়ী সাজানো একটা কম্বাইন্ড লিস্টিং -- প্রতিটা
    # রো-তে সেই লোন কোন কোন ক্যাটাগরিতে পড়ে তা (Expired/Rescheduled/Overdue/Due)
    # রেন্ডার-টাইমে যোগ হয় (_status_labels() দেখুন)।
    village_union_listing = sorted(
        filtered, key=lambda d: (_s(d.get("union")).lower(), _s(d.get("village")).lower())
    )

    return {
        "cutoff_date": cutoff_date,
        "overdue_start": overdue_start,
        "overdue_end": overdue_end,
        "overall_summary": overall_summary,
        "overall_grand": dict(grand),
        "village_union_listing": village_union_listing,
        "overdue": _sorted_flat(overdue_rows, "overdue"),
        "expired": _sorted_flat(expired_rows, "expired"),
        "rescheduled": _sorted_flat(resch_rows, "rescheduled"),
        "due": _sorted_flat(due_rows, "due"),
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

_BANGLA_RE = re.compile(r"[\u0980-\u09FF]+")


def _strip_bangla(s):
    """Comment-জাতীয় ফ্রি-টেক্সট কলামে বাংলা অক্ষর থাকলে PDF-এর ফন্ট (Helvetica)
    সেটা সাপোর্ট করে না, ফলে garbled/truncated দেখায় -- তাই সেই বাংলা অংশটুকু বাদ
    দিয়ে বাকি (ইংরেজি/সংখ্যা) অংশ রাখা হয়।"""
    return re.sub(r"\s+", " ", _BANGLA_RE.sub("", s)).strip()


def _fmt_cell(key, val):
    if val is None or str(val).strip() in ("", "None"):
        return ""
    if key in _AMOUNT_KEYS:
        try:
            return f"{float(str(val).replace(',', '')):,.0f}"
        except (TypeError, ValueError):
            return str(val)
    s = str(val)
    if key == "blank_col":
        s = _strip_bangla(s)
    return s


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
        topMargin=42 * mm, bottomMargin=12 * mm,
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


_OVERALL_SUMMARY_COLUMNS = [
    ("union", "Union"),
    ("count", "Loan Count"),
    ("balance", "Balance"),
    ("expired_count", "Expired Count"),
    ("expired_balance", "Expired Balance"),
    ("resch_count", "Rescheduled Count"),
    ("resch_balance", "Rescheduled Balance"),
    ("due_count", "Due Count"),
    ("due_balance", "Due Balance"),
]


def _build_overall_summary_table(overall_summary, overall_grand, page_w, cell_style, header_style):
    """Union-ভিত্তিক সারাংশ (Loan Count/Balance + Expired/Rescheduled/Due-এর Count+Balance)
    -- কোনো লোন-লিস্টিং না, শুধু সংখ্যার টেবিল। শেষে Grand Total সারি।"""
    def _fmt(key, val):
        if key == "union":
            return str(val)
        if key.endswith("_count") or key == "count":
            return f"{int(val):,}"
        return f"{val:,.0f}"

    header_row = [Paragraph(h, header_style) for _, h in _OVERALL_SUMMARY_COLUMNS]
    table_data = [header_row]

    for union, row in overall_summary:
        vals = [union] + [row.get(k, 0) for k, _ in _OVERALL_SUMMARY_COLUMNS[1:]]
        table_data.append([
            Paragraph(_fmt(k, v), cell_style)
            for (k, _), v in zip(_OVERALL_SUMMARY_COLUMNS, vals)
        ])

    grand_vals = ["Grand Total"] + [overall_grand.get(k, 0) for k, _ in _OVERALL_SUMMARY_COLUMNS[1:]]
    table_data.append([
        Paragraph(_fmt(k, v), header_style)
        for (k, _), v in zip(_OVERALL_SUMMARY_COLUMNS, grand_vals)
    ])

    avail_width = page_w - 16 * mm
    raw_weights = [0.13, 0.09, 0.12, 0.10, 0.12, 0.12, 0.13, 0.09, 0.10]
    total_w = sum(raw_weights)
    col_widths = [avail_width * w / total_w for w in raw_weights]

    t = Table(table_data, colWidths=col_widths, repeatRows=1)
    t.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#F2F2F2")),
        ("LINEBELOW", (0, 0), (-1, 0), 1, colors.black),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -2), [colors.white, colors.HexColor("#F2F2F2")]),
        ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#dfe6e9")),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    return t


def _build_combined_categories_table(category_data, page_w, cell_style, header_style,
                                      label_style, subtotal_para_style):
    """category_data: [(label, flat_rows), ...] -- Overdue/Expired/Rescheduled/Due,
    সবগুলো একসাথে মিলে ONE Table flowable বানায়: কলাম-হেডার রো (Sl., Loan Case, ...)
    শুধু সবার উপরে একবার (repeatRows=1 দিয়ে দরকার হলে পরের পেজে অটো রিপিট), প্রতিটা
    ক্যাটাগরির আগে merged লেবেল-রো, শেষে merged সাবটোটাল-রো। তাই একাধিক ক্যাটাগরি
    অনায়াসে একই পেজে ধরে যায়, কিন্তু কলাম-হেডার রো সারা রিপোর্টে কার্যত একবারই
    (পেজ-ওভারফ্লো ছাড়া) দেখা যায়।"""
    ncols = 1 + len(COLUMNS)
    header_row = [Paragraph("Sl.", header_style)] + [Paragraph(h, header_style) for _, h in COLUMNS]
    table_data = [header_row]
    span_cmds = []
    bg_cmds = []
    row_i = 1

    for label, rows in category_data:
        table_data.append([Paragraph(label, label_style)] + [""] * (ncols - 1))
        span_cmds.append(("SPAN", (0, row_i), (ncols - 1, row_i)))
        bg_cmds.append(("BACKGROUND", (0, row_i), (ncols - 1, row_i), colors.HexColor("#dfe6e9")))
        row_i += 1

        if not rows:
            table_data.append([Paragraph("No data found.", cell_style)] + [""] * (ncols - 1))
            span_cmds.append(("SPAN", (0, row_i), (ncols - 1, row_i)))
            row_i += 1
        else:
            for i, d in enumerate(rows, start=1):
                table_data.append(
                    [Paragraph(str(i), cell_style)] +
                    [Paragraph(_fmt_cell(k, d.get(k)), cell_style) for k, _ in COLUMNS]
                )
                row_i += 1

        total_balance = sum(_num(d.get("bal_total")) for d in rows)
        subtotal_text = (f"Total Loans: {len(rows)} &nbsp;&nbsp;|&nbsp;&nbsp; "
                          f"Total Balance: {total_balance:,.0f}")
        table_data.append([Paragraph(subtotal_text, subtotal_para_style)] + [""] * (ncols - 1))
        span_cmds.append(("SPAN", (0, row_i), (ncols - 1, row_i)))
        bg_cmds.append(("BACKGROUND", (0, row_i), (ncols - 1, row_i), colors.HexColor("#f7f7f7")))
        row_i += 1

    avail_width = page_w - 16 * mm
    total_weight = _SL_WEIGHT + sum(_COLUMN_WEIGHTS.get(k, 1.0) for k, _ in COLUMNS)
    col_widths = [avail_width * _SL_WEIGHT / total_weight] + \
                 [avail_width * _COLUMN_WEIGHTS.get(k, 1.0) / total_weight for k, _ in COLUMNS]

    t = Table(table_data, colWidths=col_widths, repeatRows=1)
    base_style = [
        ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#F2F2F2")),
        ("LINEBELOW", (0, 0), (-1, 0), 1, colors.black),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]
    t.setStyle(TableStyle(base_style + span_cmds + bg_cmds))
    return t


def _status_labels(row, cutoff_date, overdue_start, overdue_end):
    """একটা লোন Expired/Rescheduled/Overdue/Due -- এর মধ্যে কোন কোনটায় পড়ে তা টেক্সট
    হিসেবে বানায়, যেমন: "Overdue (01/09/2026 to 01/12/2026), Due"। একাধিক প্রযোজ্য
    হতে পারে (এগুলো একে অপরকে বাদ দেয় না), কোনোটাই প্রযোজ্য না হলে "Regular"।"""
    labels = []
    dt = parse_ddmmyyyy(row.get("overdue_date"))
    if dt and dt <= cutoff_date:
        labels.append(f"Expired ({cutoff_date.strftime('%d/%m/%Y')})")
    if dt and dt > cutoff_date and _num(row.get("reschedule_no")) > 0:
        labels.append("Rescheduled")
    if dt and overdue_start <= dt <= overdue_end:
        labels.append(f"Overdue ({overdue_start.strftime('%d/%m/%Y')} to {overdue_end.strftime('%d/%m/%Y')})")
    if _num(row.get("due_amount")) > 0:
        labels.append("Due")
    return ", ".join(labels) if labels else "Regular"


def _build_village_union_listing_table(rows, cutoff_date, overdue_start, overdue_end, page_w,
                                        cell_style, header_style, subtotal_para_style):
    """সিলেক্টেড স্কোপের সব লোন (কোনো ক্যাটাগরি-ফিল্টার ছাড়া) Union → Village অনুযায়ী
    সাজানো একটা কম্বাইন্ড লিস্টিং টেবিল বানায় -- মূল কলামগুলোর সাথে একটা বাড়তি
    "Status" কলাম (Expired/Rescheduled/Overdue/Due -- একাধিক প্রযোজ্য হলে সবই দেখায়)।
    প্রতিটা Union-এর শেষে একটা merged Sub Total সারি, সবশেষে Grand Total। প্রতি
    Union-এর ভেতরে Sl. নতুন করে ১ থেকে শুরু হয়। rows আগে থেকেই Union→Village
    অনুযায়ী সাজানো থাকতে হবে (build_overview_report_data()-এর village_union_listing)।
    """
    status_col = ("status", "Status")
    all_cols = list(COLUMNS) + [status_col]
    ncols = 1 + len(all_cols)
    header_row = [Paragraph("Sl.", header_style)] + [Paragraph(h, header_style) for _, h in all_cols]
    table_data = [header_row]
    span_cmds = []
    bg_cmds = []
    row_i = 1

    grand_count, grand_balance = 0, 0.0

    def _flush_union(union_label, u_rows, u_count, u_balance):
        nonlocal row_i
        for i, d in enumerate(u_rows, start=1):
            cells = [Paragraph(str(i), cell_style)]
            for k, _h in COLUMNS:
                cells.append(Paragraph(_fmt_cell(k, d.get(k)), cell_style))
            cells.append(Paragraph(_status_labels(d, cutoff_date, overdue_start, overdue_end), cell_style))
            table_data.append(cells)
            row_i += 1
        subtotal_text = (f"Sub Total ({union_label}) — Total Loans: {u_count} "
                          f"&nbsp;&nbsp;|&nbsp;&nbsp; Total Balance: {u_balance:,.0f}")
        table_data.append([Paragraph(subtotal_text, subtotal_para_style)] + [""] * (ncols - 1))
        span_cmds.append(("SPAN", (0, row_i), (ncols - 1, row_i)))
        bg_cmds.append(("BACKGROUND", (0, row_i), (ncols - 1, row_i), colors.HexColor("#f7f7f7")))
        row_i += 1

    current_union = None
    bucket = []
    for d in rows:
        union = _s(d.get("union")) or "Unknown"
        if current_union is not None and union != current_union:
            u_balance = sum(_num(x.get("bal_total")) for x in bucket)
            _flush_union(current_union, bucket, len(bucket), u_balance)
            grand_count += len(bucket)
            grand_balance += u_balance
            bucket = []
        current_union = union
        bucket.append(d)
    if bucket:
        u_balance = sum(_num(x.get("bal_total")) for x in bucket)
        _flush_union(current_union, bucket, len(bucket), u_balance)
        grand_count += len(bucket)
        grand_balance += u_balance

    grand_text = (f"Grand Total — Total Loans: {grand_count} &nbsp;&nbsp;|&nbsp;&nbsp; "
                  f"Total Balance: {grand_balance:,.0f}")
    table_data.append([Paragraph(grand_text, subtotal_para_style)] + [""] * (ncols - 1))
    span_cmds.append(("SPAN", (0, row_i), (ncols - 1, row_i)))
    bg_cmds.append(("BACKGROUND", (0, row_i), (ncols - 1, row_i), colors.HexColor("#dfe6e9")))

    avail_width = page_w - 16 * mm
    total_weight = _SL_WEIGHT + sum(_COLUMN_WEIGHTS.get(k, 1.0) for k, _ in COLUMNS) + 1.4
    col_widths = [avail_width * _SL_WEIGHT / total_weight] + \
                 [avail_width * _COLUMN_WEIGHTS.get(k, 1.0) / total_weight for k, _ in COLUMNS] + \
                 [avail_width * 1.4 / total_weight]

    t = Table(table_data, colWidths=col_widths, repeatRows=1)
    base_style = [
        ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#F2F2F2")),
        ("LINEBELOW", (0, 0), (-1, 0), 1, colors.black),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]
    t.setStyle(TableStyle(base_style + span_cmds + bg_cmds))
    return t


def generate_overview_report_pdf(overview_data, out_path, branch_name,
                                  bank_name=BANK_NAME_DEFAULT, logo_path=None,
                                  show_village_union_listing=False):
    """overview_data: build_overview_report_data()-এর রেজাল্ট। পেজের উপরের টাইটেল
    সবসময় শুধু "Overview Report" (তার নিচে cutoff/overdue-range তথ্যের একটা লাইন)।
    তারপর "Overall Loan Information" (Union-ভিত্তিক সারাংশ টেবিল, কোনো লোন-লিস্টিং
    না), তারপর (show_village_union_listing=True হলে) "Village/Union Wise" -- সিলেক্টেড
    স্কোপের সব লোন একসাথে Union→Village অনুযায়ী সাজানো, প্রতিটা রো-তে Expired/
    Rescheduled/Overdue/Due স্ট্যাটাস দেখানো, Union-ভিত্তিক Sub Total + Grand Total সহ।
    সবশেষে Overdue → Expired → Rescheduled → Due -- এই ৪টা মিলে ONE Table-এ, তাই
    কলাম-হেডার রো (Sl., Loan Case, ...) কার্যত সারা রিপোর্টে একবারই থাকে -- একাধিক
    ক্যাটাগরি অনায়াসে একই পেজে ধরে যায়।"""
    page_w, page_h = landscape(A4)
    default_logo = os.path.join(os.path.dirname(__file__), "logo.png")
    if logo_path is None and os.path.exists(default_logo):
        logo_path = default_logo

    doc = SimpleDocTemplate(
        out_path, pagesize=landscape(A4),
        leftMargin=8 * mm, rightMargin=8 * mm,
        topMargin=42 * mm, bottomMargin=12 * mm,
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
    label_style = ParagraphStyle(
        "catlabel", parent=styles["Normal"], fontSize=10, fontName=FONT_BOLD, leading=12,
    )
    subtotal_para_style = ParagraphStyle(
        "subtotal_inline", parent=styles["Normal"], fontSize=8, fontName=FONT_BOLD, leading=10,
    )
    info_style = ParagraphStyle("info", parent=styles["Normal"], fontSize=9,
                                 fontName=FONT_REGULAR, spaceAfter=8)

    cutoff = overview_data.get("cutoff_date")
    ostart = overview_data.get("overdue_start")
    oend = overview_data.get("overdue_end")

    elements = [Paragraph(
        f"Cut-off Date (Expired/Rescheduled): "
        f"{cutoff.strftime('%d/%m/%Y') if cutoff else '-'} &nbsp;&nbsp;|&nbsp;&nbsp; "
        f"Overdue Range: {ostart.strftime('%d/%m/%Y') if ostart else '-'} to "
        f"{oend.strftime('%d/%m/%Y') if oend else '-'}",
        info_style,
    )]

    elements.append(Paragraph("Overall Loan Information", section_style))
    elements.append(_build_overall_summary_table(
        overview_data.get("overall_summary") or [], overview_data.get("overall_grand") or {},
        page_w, cell_style, header_style,
    ))
    elements.append(Spacer(1, 10))

    if show_village_union_listing:
        elements.append(Paragraph("Village/Union Wise", section_style))
        elements.append(_build_village_union_listing_table(
            overview_data.get("village_union_listing") or [],
            overview_data.get("cutoff_date"), overview_data.get("overdue_start"),
            overview_data.get("overdue_end"), page_w, cell_style, header_style, subtotal_para_style,
        ))
        elements.append(Spacer(1, 10))

    category_data = [(label, overview_data.get(key) or []) for key, label in OVERVIEW_SECTIONS]
    elements.append(_build_combined_categories_table(
        category_data, page_w, cell_style, header_style, label_style, subtotal_para_style,
    ))

    print_dt_str = datetime.now(BD_TZ).strftime("%d/%m/%Y %I:%M %p")
    header_fn = lambda c, d: _draw_header(c, d, bank_name, branch_name, logo_path,
                                           "Overview Report", print_dt_str)
    doc.build(elements, onFirstPage=header_fn, onLaterPages=header_fn)
    return out_path


def generate_overview_excel(overview_data, out_path, show_village_union_listing=False):
    """overview_data: build_overview_report_data()-এর রেজাল্ট। প্রথমে "Overall Loan
    Information" (Union-ভিত্তিক সারাংশ, Grand Total সহ) ব্লক, তারপর
    (show_village_union_listing=True হলে) "Village/Union Wise" -- সব লোন Union→Village
    অনুযায়ী সাজানো + Status কলাম (Expired/Rescheduled/Overdue/Due) + Union-ভিত্তিক
    Sub Total ও Grand Total -- তারপর Overdue → Expired → Rescheduled → Due -- প্রতিটা
    ফ্ল্যাট (Union-গ্রুপ ছাড়া, একটা আলাদা Union কলাম দিয়ে চেনা যাবে) ব্লক হিসেবে,
    প্রতিটার আগে লেবেল রো + হেডার রো। ব্যাংকের নাম/লোগো/টাইটেল নেই।"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Overview Report"

    r = 1
    ws.cell(row=r, column=1, value="Overall Loan Information").font = Font(bold=True, size=13)
    r += 1
    for idx, (_key, h) in enumerate(_OVERALL_SUMMARY_COLUMNS, start=1):
        ws.cell(row=r, column=idx, value=h).font = Font(bold=True)
    r += 1
    for union, row in overview_data.get("overall_summary") or []:
        vals = [union] + [row.get(k, 0) for k, _ in _OVERALL_SUMMARY_COLUMNS[1:]]
        for idx, v in enumerate(vals, start=1):
            ws.cell(row=r, column=idx, value=v)
        r += 1
    grand = overview_data.get("overall_grand") or {}
    grand_vals = ["Grand Total"] + [grand.get(k, 0) for k, _ in _OVERALL_SUMMARY_COLUMNS[1:]]
    for idx, v in enumerate(grand_vals, start=1):
        ws.cell(row=r, column=idx, value=v).font = Font(bold=True)
    r += 3

    if show_village_union_listing:
        cutoff_date = overview_data.get("cutoff_date")
        overdue_start = overview_data.get("overdue_start")
        overdue_end = overview_data.get("overdue_end")
        listing = overview_data.get("village_union_listing") or []

        ws.cell(row=r, column=1, value="Village/Union Wise").font = Font(bold=True, size=13)
        r += 1
        vu_header = ["Sl."] + [h for _, h in COLUMNS] + ["Status"]
        for idx, h in enumerate(vu_header, start=1):
            ws.cell(row=r, column=idx, value=h).font = Font(bold=True)
        r += 1

        grand_count, grand_balance = 0, 0.0
        current_union, bucket = None, []

        def _flush(union_label, u_rows):
            nonlocal r
            for i, d in enumerate(u_rows, start=1):
                row_vals = ([i] + [_excel_cell(k, d.get(k)) for k, _ in COLUMNS] +
                            [_status_labels(d, cutoff_date, overdue_start, overdue_end)])
                for idx, v in enumerate(row_vals, start=1):
                    ws.cell(row=r, column=idx, value=v)
                r += 1
            u_balance = sum(_num(x.get("bal_total")) for x in u_rows)
            sub_cell = ws.cell(row=r, column=1,
                                value=f"Sub Total ({union_label}) — Loans: {len(u_rows)}, Balance: {u_balance:,.0f}")
            sub_cell.font = Font(bold=True)
            r += 1
            return len(u_rows), u_balance

        for d in listing:
            union = _s(d.get("union")) or "Unknown"
            if current_union is not None and union != current_union:
                c, b = _flush(current_union, bucket)
                grand_count += c
                grand_balance += b
                bucket = []
            current_union = union
            bucket.append(d)
        if bucket:
            c, b = _flush(current_union, bucket)
            grand_count += c
            grand_balance += b

        grand_cell = ws.cell(row=r, column=1,
                              value=f"Grand Total — Loans: {grand_count}, Balance: {grand_balance:,.0f}")
        grand_cell.font = Font(bold=True)
        r += 3

    for key, label in OVERVIEW_SECTIONS:
        flat_rows = overview_data.get(key) or []
        ws.cell(row=r, column=1, value=label).font = Font(bold=True, size=13)
        r += 1

        header = ["Sl."] + [h for _, h in COLUMNS] + ["Union"]
        for idx, h in enumerate(header, start=1):
            ws.cell(row=r, column=idx, value=h).font = Font(bold=True)
        r += 1

        for i, d in enumerate(flat_rows, start=1):
            row_vals = [i] + [_excel_cell(k, d.get(k)) for k, _ in COLUMNS] + [_s(d.get("union"))]
            for idx, v in enumerate(row_vals, start=1):
                ws.cell(row=r, column=idx, value=v)
            r += 1
        r += 2  # পরের সেকশনের আগে ফাঁকা গ্যাপ

    widths = [6] + [max(10, len(h) + 2) for _, h in COLUMNS] + [14]
    for idx, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = w

    wb.save(out_path)
    return out_path
