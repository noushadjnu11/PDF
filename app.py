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

import pandas as pd
import streamlit as st

from pdf_processor import (
    parse_borrower_list_pdf,
    parse_loan_balance_pdf,
    build_prefix_map,
    merge_sources,
    write_excel,
)

st.set_page_config(page_title="Loan PDF Merge Tool", page_icon="🏦", layout="centered")

st.title("🏦 Karmasangsthan Bank — Loan PDF Merge Tool")
st.caption("Borrower List PDF এবং Loan Balance PDF আপলোড করুন — একটা মার্জড Excel ফাইল পাবেন।")

# ---------------------------------------------------------------------------
# সেশন স্টেট ইনিশিয়ালাইজ
# ---------------------------------------------------------------------------
for key, default in [
    ("borrower_rows", None), ("balance_rows", None),
    ("programs", None), ("scanned", False),
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

    df = pd.DataFrame([
        {"প্রোগ্রাম কোড": code, "প্রোগ্রামের নাম": name, "Prefix": auto_prefix.get(code, code)}
        for code, name in sorted(programs.items())
    ])

    edited_df = st.data_editor(
        df,
        column_config={
            "প্রোগ্রাম কোড": st.column_config.TextColumn(disabled=True),
            "প্রোগ্রামের নাম": st.column_config.TextColumn(disabled=True),
            "Prefix": st.column_config.TextColumn(help="এই প্রোগ্রামের লোন কেসগুলোর সামনে এই prefix বসবে"),
        },
        hide_index=True,
        use_container_width=True,
        key="prefix_editor",
    )

    prefix_overrides = dict(zip(edited_df["প্রোগ্রাম কোড"], edited_df["Prefix"]))

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
else:
    st.info("প্রথমে দুইটা PDF আপলোড করে 'PDF স্ক্যান করুন' বাটনে ক্লিক করুন।")
