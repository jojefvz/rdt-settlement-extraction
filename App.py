"""
Jim Daws Trucking — Settlement Processor
Streamlit UI wrapper around process_settlements.py
"""

import io
import tempfile
import os
import streamlit as st
from process_settlements import extract_settlements, settlement_to_rows, write_excel

# ── Page config ───────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Jim Daws Trucking — Settlement Processor",
    page_icon="🚚",
    layout="centered",
)

# ── Header ────────────────────────────────────────────────────────────────────

st.title("🚚 Jim Daws Trucking")
st.subheader("Settlement Processor")
st.caption("Upload a weekly driver settlement PDF to generate a QuickBooks-ready Excel file.")
st.divider()

# ── File upload ───────────────────────────────────────────────────────────────

uploaded_file = st.file_uploader(
    "Upload Settlement PDF",
    type="pdf",
    help="Select the weekly settlement PDF exported from your system."
)

if uploaded_file:
    st.success(f"📄 **{uploaded_file.name}** uploaded successfully.")
    st.divider()

    if st.button("⚙️ Process Settlements", type="primary", use_container_width=True):

        with st.spinner("Processing settlements..."):

            # Write uploaded PDF to a temp file so pdfplumber can read it
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
                tmp_pdf.write(uploaded_file.read())
                tmp_pdf_path = tmp_pdf.name

            try:
                # Extract settlements
                settlements = extract_settlements(tmp_pdf_path)

                # Build Excel in memory
                output_name = uploaded_file.name.replace(".pdf", "_QB.xlsx")
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_xlsx:
                    tmp_xlsx_path = tmp_xlsx.name

                write_excel(settlements, tmp_xlsx_path)

                with open(tmp_xlsx_path, "rb") as f:
                    excel_bytes = f.read()

            finally:
                os.unlink(tmp_pdf_path)
                if os.path.exists(tmp_xlsx_path):
                    os.unlink(tmp_xlsx_path)

        # ── Summary ───────────────────────────────────────────────────────────
        st.divider()
        st.success("✅ Processing complete!")

        total_rows = sum(len(settlement_to_rows(s)) for s in settlements)
        total_skipped = sum(
            len(s['skipped_other_pay']) + len(s['skipped_total_settlement']) + len(s['skipped_reserves'])
            for s in settlements
        )
        unmapped = [
            s for s in settlements
            if (s['truck_note'] is not None and s['truck'] not in
                __import__('process_settlements').NOTES_RECEIVABLE) or
               (s['maint_fund'] is not None and s['truck'] not in
                __import__('process_settlements').MAINT_FUND_ACCOUNTS)
        ]

        col1, col2, col3 = st.columns(3)
        col1.metric("Drivers Processed", len(settlements))
        col2.metric("QB Rows Written", total_rows)
        col3.metric("Skipped Lines", total_skipped)

        if unmapped:
            st.warning(
                f"⚠️ **{len(unmapped)} driver(s) have unmapped accounts** — "
                f"check the Skipped Lines tab in the Excel file:\n" +
                "\n".join(f"- {s['driver_name']} (TRK:{s['truck']})" for s in unmapped)
            )

        st.divider()

        # ── Driver breakdown ──────────────────────────────────────────────────
        with st.expander("📋 Driver breakdown", expanded=False):
            for s in settlements:
                inv = s['invoice_number'] or 'NO INV#'
                trk = s['truck'] or 'NO TRK'
                skipped = (
                    len(s['skipped_other_pay']) +
                    len(s['skipped_total_settlement']) +
                    len(s['skipped_reserves'])
                )
                flag = " ⚠️" if (
                    (s['truck_note'] is not None and s['truck'] not in
                     __import__('process_settlements').NOTES_RECEIVABLE) or
                    (s['maint_fund'] is not None and s['truck'] not in
                     __import__('process_settlements').MAINT_FUND_ACCOUNTS)
                ) else ""
                st.markdown(
                    f"**{s['driver_name']}**{flag} &nbsp;|&nbsp; "
                    f"TRK: `{trk}` &nbsp;|&nbsp; "
                    f"INV: `{inv}` &nbsp;|&nbsp; "
                    f"Gross: `${s['gross_pay']:,.2f}` &nbsp;|&nbsp; "
                    f"Skipped: `{skipped}`"
                )

        # ── Download button ───────────────────────────────────────────────────
        st.download_button(
            label="⬇️ Download Excel File",
            data=excel_bytes,
            file_name=output_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
        )
