# -*- coding: utf-8 -*-
"""
Karmasangsthan Bank — Loan PDF Merge Tool (Streamlit App)
============================================================
রান করার নিয়ম (README_বাংলা.md ফাইলেও বিস্তারিত আছে):
    pip install -r requirements.txt
    streamlit run app.py
"""

import os
import tempfile
from datetime import date

import pandas as pd
import streamlit as st

from pdf_processor import (
    parse_borrower_list_pdf,
    parse_loan_balance_pdf,
    build_prefix_map,
    merge_sources,
    write_excel,
)
import report_generator as rg

st.set_page_config(page_title="Loan PDF Merge Tool", page_icon="🏦", layout="centered")

st.title("🏦 Karmasangsthan Bank — Loan PDF Merge Tool")
st.caption("Borrower List PDF এবং Loan Balance PDF আপলোড করুন — একটা মার্জড Excel ফাইল পাবেন।")

# ---------------------------------------------------------------------------
# সেশন স্টেট ইনিশিয়ালাইজ
# ---------------------------------------------------------------------------
for key, default in [
    ("borrower_rows", None), ("balance_rows", None),
    ("programs", None), ("scanned", False), ("merged_xlsx_path", None),
]:
    if key not in st.session_state:
        st.session_state[key] = default


# ---------------------------------------------------------------------------
# ধাপ ১: PDF আপলোড
# ---------------------------------------------------------------------------
st.header("ধাপ ১ — PDF আপলোড করুন")

col1, col2 = st.columns(2)
with col1:
    borrower_file = st.file_uploader("Borrower List PDF", type="pdf", key="borrower_upload")
with col2:
    balance_file = st.file_uploader("Loan Balance PDF", type="pdf", key="balance_upload")


def save_temp(uploaded_file):
    suffix = ".pdf"
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp.write(uploaded_file.getbuffer())
    tmp.close()
    return tmp.name


# ---------------------------------------------------------------------------
# ধাপ ২: PDF স্ক্যান করে Loan Program-এর তালিকা বের করা (prefix দেখানো/এডিট করার জন্য)
# ---------------------------------------------------------------------------
st.header("ধাপ ২ — Loan Program ও Prefix যাচাই করুন")

scan_btn = st.button("🔍 PDF স্ক্যান করুন", type="primary",
                      disabled=not (borrower_file and balance_file))

if scan_btn:
    with st.spinner("PDF পড়া হচ্ছে... (কিছুক্ষণ সময় লাগতে পারে)"):
        b_path = save_temp(borrower_file)
        l_path = save_temp(balance_file)
        try:
            borrower_rows, borrower_programs = parse_borrower_list_pdf(b_path)
            balance_rows, balance_programs = parse_loan_balance_pdf(l_path)
        finally:
            os.unlink(b_path)
            os.unlink(l_path)

        all_programs = {**balance_programs, **borrower_programs}
        auto_prefix_map = build_prefix_map(all_programs)

        st.session_state["borrower_rows"] = borrower_rows
        st.session_state["balance_rows"] = balance_rows
        st.session_state["programs"] = all_programs
        st.session_state["auto_prefix"] = auto_prefix_map
        st.session_state["scanned"] = True

    st.success(f"স্ক্যান সম্পন্ন। Borrower List-এ {len(borrower_rows)}টা এবং "
               f"Loan Balance-এ {len(balance_rows)}টা লোন কেস পাওয়া গেছে।")

if st.session_state["scanned"]:
    programs = st.session_state["programs"]
    auto_prefix = st.session_state["auto_prefix"]

    st.write("নিচের তালিকায় প্রতিটা Loan Program-এর জন্য স্বয়ংক্রিয়ভাবে একটা Prefix বসানো হয়েছে। "
             "প্রয়োজনে **Prefix** কলামে ক্লিক করে নিজের মতো বদলে দিতে পারেন।")

    st.caption("আগে ডাউনলোড করা Prefix তালিকার CSV থাকলে এখানে আপলোড করে দিলে সেটা থেকেই বসে যাবে।")
    prefix_csv_file = st.file_uploader("Prefix তালিকার CSV আপলোড করুন (ঐচ্ছিক)",
                                        type="csv", key="prefix_csv_upload")

    uploaded_prefix_map = {}
    csv_ok = False
    if prefix_csv_file is not None:
        try:
            csv_df = pd.read_csv(prefix_csv_file, dtype=str).fillna("")
            code_col, prefix_col = "প্রোগ্রাম কোড", "Prefix"
            if code_col in csv_df.columns and prefix_col in csv_df.columns:
                uploaded_prefix_map = dict(zip(csv_df[code_col], csv_df[prefix_col]))
                csv_ok = True
            else:
                st.error("CSV ফরম্যাট সঠিক নয় — 'প্রোগ্রাম কোড' এবং 'Prefix' নামে কলাম থাকা দরকার।")
        except Exception as e:
            st.error(f"CSV ফাইল পড়তে সমস্যা হয়েছে: {e}")

    missing_programs = []
    if csv_ok:
        missing_programs = [
            (code, name) for code, name in sorted(programs.items())
            if code not in uploaded_prefix_map or not str(uploaded_prefix_map[code]).strip()
        ]
        if missing_programs:
            missing_list_html = "".join(
                f"<li><b>{code}</b> — {name}</li>" for code, name in missing_programs
            )
            st.markdown(
                "<div style='color:#D32F2F; font-weight:bold;'>⚠️ আপলোড করা CSV-তে নিচের "
                "প্রোগ্রামগুলোর Prefix পাওয়া যায়নি — এগুলো নিচের তালিকায় লাল চিহ্নিত, "
                "নিজে থেকে Prefix বসিয়ে ঠিক করে দিন:</div>"
                f"<ul style='color:#D32F2F;'>{missing_list_html}</ul>",
                unsafe_allow_html=True,
            )
        else:
            st.success("✅ CSV থেকে সব প্রোগ্রামের Prefix পাওয়া গেছে।")

    missing_codes = {code for code, _ in missing_programs}

    df = pd.DataFrame([
        {
            "স্ট্যাটাস": "⚠️ মিসিং" if code in missing_codes else "",
            "প্রোগ্রাম কোড": code,
            "প্রোগ্রামের নাম": name,
            "Prefix": uploaded_prefix_map.get(code) or auto_prefix.get(code, code),
        }
        for code, name in sorted(programs.items())
    ])

    edited_df = st.data_editor(
        df,
        column_config={
            "স্ট্যাটাস": st.column_config.TextColumn(disabled=True, help="CSV আপলোডের পর মিসিং প্রোগ্রাম এখানে দেখাবে"),
            "প্রোগ্রাম কোড": st.column_config.TextColumn(disabled=True),
            "প্রোগ্রামের নাম": st.column_config.TextColumn(disabled=True),
            "Prefix": st.column_config.TextColumn(help="এই প্রোগ্রামের লোন কেসগুলোর সামনে এই prefix বসবে"),
        },
        hide_index=True,
        use_container_width=True,
        key="prefix_editor",
    )

    prefix_overrides = dict(zip(edited_df["প্রোগ্রাম কোড"], edited_df["Prefix"]))

    csv_export = edited_df[["প্রোগ্রাম কোড", "প্রোগ্রামের নাম", "Prefix"]].to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "⬇️ Prefix তালিকা CSV আকারে ডাউনলোড করুন",
        data=csv_export,
        file_name="prefix_list.csv",
        mime="text/csv",
    )

    # ---------------------------------------------------------------------------
    # ধাপ ৩: মার্জ করে Excel বানানো
    # ---------------------------------------------------------------------------
    st.header("ধাপ ৩ — Merge করে Excel বানান")

    if st.button("✅ Merge করে Excel তৈরি করুন", type="primary"):
        with st.spinner("মার্জ করা হচ্ছে..."):
            final_prefix_map = build_prefix_map(programs, overrides=prefix_overrides)
            merged = merge_sources(
                st.session_state["borrower_rows"],
                st.session_state["balance_rows"],
                final_prefix_map,
            )
            out_path = os.path.join(tempfile.gettempdir(), "merged_loan_report.xlsx")
            stats = write_excel(merged, out_path)

        st.success("✅ সম্পন্ন! নিচে ডাউনলোড করুন।")

        c1, c2 = st.columns(2)
        c1.metric("মোট রো", stats["total_rows"])
        c2.metric("⚠ ম্যানুয়াল যাচাই দরকার এমন রো", stats["needs_review"])

        if stats["needs_review"] > 0:
            st.warning(
                f"{stats['needs_review']}টা রো-তে নাম একাধিক লাইনে ভাগ হয়ে যাওয়ায় "
                f"Father/Spouse/Village সঠিকভাবে আলাদা করা যায়নি। এই রো-গুলো Excel-এ "
                f"হলুদ রঙে চিহ্নিত এবং 'নোট' কলামে কারণ লেখা আছে — একবার চোখ বুলিয়ে "
                f"দেখে নেবেন।"
            )

        with open(out_path, "rb") as f:
            st.download_button(
                "⬇️ Excel ফাইল ডাউনলোড করুন",
                data=f.read(),
                file_name="merged_loan_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )

        st.session_state["merged_xlsx_path"] = out_path
else:
    st.info("প্রথমে দুইটা PDF আপলোড করে 'PDF স্ক্যান করুন' বাটনে ক্লিক করুন।")


# ---------------------------------------------------------------------------
# ধাপ ৪: Excel থেকে ফিল্টার করে A4 Landscape PDF রিপোর্ট বানানো
# ---------------------------------------------------------------------------
st.header("ধাপ ৪ — PDF রিপোর্ট জেনারেট করুন")

report_source = st.radio(
    "কোন Excel থেকে রিপোর্ট বানাবেন?",
    ["এইমাত্র মার্জ করা Excel (উপরে)", "অন্য একটা Excel আপলোড করুন (সংশোধিত ফাইল)"],
    horizontal=True,
)

report_xlsx_path = None
if report_source == "এইমাত্র মার্জ করা Excel (উপরে)":
    if st.session_state.get("merged_xlsx_path"):
        report_xlsx_path = st.session_state["merged_xlsx_path"]
    else:
        st.info("এখনো কোনো Excel মার্জ করা হয়নি — আগে ধাপ ৩ সম্পন্ন করুন, অথবা ডানের অপশনে "
                 "নিজের Excel আপলোড করুন।")
else:
    uploaded_xlsx = st.file_uploader(
        "সংশোধিত Excel ফাইল আপলোড করুন (Father/Village/Union ঠিক করা)",
        type=["xlsx"], key="report_xlsx_upload",
    )
    if uploaded_xlsx:
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        tmp.write(uploaded_xlsx.getbuffer())
        tmp.close()
        report_xlsx_path = tmp.name

if report_xlsx_path:
    try:
        report_rows = rg.read_merged_excel(report_xlsx_path)
    except Exception as e:
        st.error(f"Excel পড়তে সমস্যা হয়েছে: {e}")
        report_rows = None

    if report_rows is not None:
        st.success(f"{len(report_rows)}টা রো পাওয়া গেছে।")

        branch_name = st.text_input("ব্রাঞ্চের নাম", value="")

        report_type = st.selectbox(
            "রিপোর্টের ধরন বাছাই করুন",
            [
                "Overdue Loan (নির্দিষ্ট তারিখ পর্যন্ত, oldest→newest)",
                "Expired Loan List (নির্দিষ্ট তারিখের আগের সব)",
                "Rescheduled Loan (নির্দিষ্ট তারিখের পরের + Reschedule No. > 0)",
                "Union/Village সাজানো + সাবটোটাল রিপোর্ট",
                "Due Amount Report (যেসব রো-তে Due Amount আছে)",
            ],
        )

        all_unions = rg.get_unions(report_rows)


        def union_village_picker(unions_key, villages_key, help_optional=True):
            """একটা Union multiselect + তার নিচে (বাছাই করা Union-এর মধ্যে থেকে)
            ঐচ্ছিক Village multiselect দেখায়। রিটার্ন করে (selected_unions, selected_villages)।"""
            label = "Union বেছে নিন (এক বা একাধিক)" if not help_optional else \
                "Union বেছে নিন (ঐচ্ছিক — খালি রাখলে সব Union দেখাবে)"
            sel_unions = st.multiselect(label, all_unions, key=unions_key)
            village_options = rg.get_villages_for_unions(report_rows, sel_unions or None)
            sel_villages = st.multiselect(
                "Village বেছে নিন (ঐচ্ছিক — খালি রাখলে বাছাই করা Union-এর সব Village দেখাবে)",
                village_options, key=villages_key,
            )
            return sel_unions, sel_villages

        gen_btn = False
        pdf_rows = None
        grouped = False
        title_text = ""
        summary = None
        out_filename = "loan_report.pdf"

        if report_type.startswith("Overdue Loan"):
            col1, col2 = st.columns(2)
            start = col1.date_input("শুরুর তারিখ", value=date(2026, 1, 1), format="DD/MM/YYYY")
            end = col2.date_input("শেষের তারিখ", value=date(2026, 6, 30), format="DD/MM/YYYY")

            overdue_unions, overdue_villages = union_village_picker("overdue_unions", "overdue_villages")

            if start > end:
                st.error("শুরুর তারিখ শেষের তারিখের পরে হতে পারবে না।")
                gen_btn = False
            else:
                gen_btn = st.button("📄 PDF রিপোর্ট বানান", type="primary", key="gen_overdue")

            if gen_btn:
                pdf_rows = rg.filter_overdue(report_rows, start, end,
                                              overdue_unions or None, overdue_villages or None)
                title_text = f"Overdue Loan from {start.strftime('%d/%m/%Y')} to {end.strftime('%d/%m/%Y')}"
                if overdue_unions:
                    title_text += f" — {', '.join(overdue_unions)}"
                if overdue_villages:
                    title_text += f" ({', '.join(overdue_villages)})"
                out_filename = rg.build_output_filename(
                    "Overdue", unions=overdue_unions, villages=overdue_villages, start=start, end=end,
                )

        elif report_type.startswith("Expired"):
            before = st.date_input("Expired Loan List up to", value=date.today(), format="DD/MM/YYYY",
                                    key="expired_date")
            expired_unions, expired_villages = union_village_picker("expired_unions", "expired_villages")
            gen_btn = st.button("📄 PDF রিপোর্ট বানান", type="primary", key="gen_expired")
            if gen_btn:
                pdf_rows = rg.filter_expired(report_rows, before, expired_unions or None,
                                              expired_villages or None)
                title_text = f"Expired Loan List up to {before.strftime('%d/%m/%Y')}"
                if expired_unions:
                    title_text += f" — {', '.join(expired_unions)}"
                if expired_villages:
                    title_text += f" ({', '.join(expired_villages)})"
                summary = {
                    "label": "Balance",
                    "count": len(pdf_rows),
                    "value": sum(rg._num(d.get("bal_total")) for d in pdf_rows),
                }
                out_filename = rg.build_output_filename("Expired", unions=expired_unions,
                                                          villages=expired_villages, single_date=before)

        elif report_type.startswith("Rescheduled"):
            after = st.date_input("Rescheduled Loan up to", value=date.today(), format="DD/MM/YYYY",
                                   key="resch_date")
            resch_unions, resch_villages = union_village_picker("resch_unions", "resch_villages")
            gen_btn = st.button("📄 PDF রিপোর্ট বানান", type="primary", key="gen_resch")
            if gen_btn:
                pdf_rows = rg.filter_rescheduled(report_rows, after, resch_unions or None,
                                                  resch_villages or None)
                title_text = f"Rescheduled Loan up to {after.strftime('%d/%m/%Y')}"
                if resch_unions:
                    title_text += f" — {', '.join(resch_unions)}"
                if resch_villages:
                    title_text += f" ({', '.join(resch_villages)})"
                summary = {
                    "label": "Balance",
                    "count": len(pdf_rows),
                    "value": sum(rg._num(d.get("bal_total")) for d in pdf_rows),
                }
                out_filename = rg.build_output_filename("Rescheduled", unions=resch_unions,
                                                          villages=resch_villages, single_date=after)

        elif report_type.startswith("Union/Village"):
            grouped_unions, grouped_villages = union_village_picker("grouped_unions", "grouped_villages")
            gen_btn = st.button("📄 PDF রিপোর্ট বানান", type="primary", key="gen_grouped")
            if gen_btn:
                pdf_rows = rg.group_by_union_village(report_rows, grouped_unions or None,
                                                       grouped_villages or None)
                title_text = "Union / Village Wise Loan Report"
                if grouped_unions:
                    title_text += f" — {', '.join(grouped_unions)}"
                if grouped_villages:
                    title_text += f" ({', '.join(grouped_villages)})"
                grouped = True
                total_count = sum(g[2]["count"] for g in pdf_rows)
                total_balance = sum(g[2]["balance"] for g in pdf_rows)
                summary = {"label": "Balance", "count": total_count, "value": total_balance}
                out_filename = rg.build_output_filename("UnionVillage", unions=grouped_unions,
                                                          villages=grouped_villages)

        else:  # Due Amount Report
            due_unions, due_villages = union_village_picker("due_unions", "due_villages")
            gen_btn = st.button("📄 PDF রিপোর্ট বানান", type="primary", key="gen_due")
            if gen_btn:
                pdf_rows = rg.filter_due_amount(report_rows, due_unions or None, due_villages or None)
                title_text = "Due Amount Report"
                if due_unions:
                    title_text += f" — {', '.join(due_unions)}"
                if due_villages:
                    title_text += f" ({', '.join(due_villages)})"
                summary = {
                    "label": "Due Amount",
                    "count": len(pdf_rows),
                    "value": sum(rg._num(d.get("due_amount")) for d in pdf_rows),
                }
                out_filename = rg.build_output_filename("DueAmount", unions=due_unions, villages=due_villages)

        if gen_btn and pdf_rows is not None:
            if not branch_name.strip():
                st.warning("ব্রাঞ্চের নাম দেওয়া হয়নি — রিপোর্টে খালি দেখাবে। তারপরও এগিয়ে যাচ্ছি।")
            with st.spinner("PDF তৈরি হচ্ছে..."):
                out_pdf = os.path.join(tempfile.gettempdir(), out_filename)
                rg.generate_report_pdf(
                    pdf_rows, out_pdf, branch_name=branch_name or "-",
                    title_text=title_text, grouped=grouped, summary=summary,
                )
            row_count = sum(len(g[1]) for g in pdf_rows) if grouped else len(pdf_rows)
            st.success(f"✅ PDF রেডি ({row_count}টা রো)।")
            with open(out_pdf, "rb") as f:
                st.download_button(
                    "⬇️ PDF রিপোর্ট ডাউনলোড করুন",
                    data=f.read(),
                    file_name=out_filename,
                    mime="application/pdf",
                    type="primary",
                )
