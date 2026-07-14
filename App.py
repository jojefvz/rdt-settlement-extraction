"""
Jim Daws Trucking — Settlement Processor
Streamlit UI wrapper around process_settlements.py
"""

import tempfile
import os
import streamlit as st
import pandas as pd
from streamlit_gsheets import GSheetsConnection
from process_settlements import extract_settlements, settlement_to_rows, write_excel, DRIVERS, load_drivers_from_sheets

DRIVERS = load_drivers_from_sheets()

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

# ── Driver Management Section ─────────────────────────────────────────────────
with st.expander("👥 Manage Driver Mappings (Google Sheets)", expanded=False):
    # Connect to your Google Sheet instance
    conn = st.connection("gsheets", type=GSheetsConnection)
    
    try:
        # Read the spreadsheet using a 1-minute cache so changes reflect quickly
        df_drivers = conn.read(worksheet="Drivers", ttl="1m")
    except Exception as e:
        st.error(f"Could not connect to Google Sheets: {e}")
        df_drivers = pd.DataFrame()

    if not df_drivers.empty:
        # 1. Display Current Configuration View
        st.dataframe(
            df_drivers[["Driver Code", "Driver Name", "Truck"]], 
            use_container_width=True, 
            hide_index=True
        )
        
        tab_add, tab_remove = st.tabs(["➕ Add Driver", "❌ Remove Driver"])
        
        # 2. Add New Entry Workflow Form
        with tab_add:
            with st.form("add_driver_form", clear_on_submit=True):
                col1, col2 = st.columns(2)
                new_code = col1.text_input("Driver Code (e.g., ROLBRE)").upper().strip()
                new_name = col2.text_input("Driver Name")
                
                col3, col4 = st.columns(2)
                new_truck = col3.text_input("Truck Number")
                new_class = col4.text_input("Driver Class")
                
                st.markdown("**Asset Account Mappings (Optional)**")
                col5, col6 = st.columns(2)
                nr_desc = col5.text_input("Notes Rec Description")
                nr_acct = col6.text_input("Notes Rec Account String")
                
                col7, col8 = st.columns(2)
                mf_desc = col7.text_input("Maint Fund Description")
                mf_acct = col8.text_input("Maint Fund Account String")
                
                submit_add = st.form_submit_button("Save New Driver", type="primary")
                
                if submit_add:
                    if not new_code or not new_name:
                        st.error("Driver Code and Driver Name are required!")
                    elif new_code in df_drivers["Driver Code"].values:
                        st.error(f"Driver Code '{new_code}' already exists!")
                    else:
                        # Append the new record to our dataframe structure
                        new_row = pd.DataFrame([{
                            "Driver Code": new_code,
                            "Driver Name": new_name,
                            "Truck": new_truck,
                            "Class": new_class,
                            "Notes Rec Desc": nr_desc,
                            "Notes Rec Account": nr_acct,
                            "Maint Fund Desc": mf_desc,
                            "Maint Fund Account": mf_acct
                        }])
                        updated_df = pd.concat([df_drivers, new_row], ignore_index=True)
                        
                        # Push updates back to Google Sheet
                        conn.update(worksheet="Drivers", data=updated_df)
                        st.cache_data.clear() # Wipe the app cache to force an immediate query update
                        st.success(f"Successfully added driver {new_name}!")
                        st.rerun()

        # 3. Delete Entry Workflow Form
        with tab_remove:
            with st.form("remove_driver_form"):
                driver_to_remove = st.selectbox(
                    "Select Driver to Remove",
                    options=df_drivers["Driver Code"].values,
                    format_func=lambda x: f"{x} - {df_drivers[df_drivers['Driver Code'] == x]['Driver Name'].values[0]}"
                )
                submit_remove = st.form_submit_button("Delete Driver", type="secondary")
                
                if submit_remove:
                    # Filter out the selected driver row
                    updated_df = df_drivers[df_drivers["Driver Code"] != driver_to_remove]
                    
                    # Push trimmed dataset back to Google Sheet
                    conn.update(worksheet="Drivers", data=updated_df)
                    st.cache_data.clear() # Wipe cache to load the update instantly
                    st.success(f"Successfully removed driver {driver_to_remove}!")
                    st.rerun()
    else:
        st.warning("No tracking data found. Please check your 'Drivers' worksheet layout headers.")

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

            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
                tmp_pdf.write(uploaded_file.read())
                tmp_pdf_path = tmp_pdf.name

            tmp_xlsx_path = None
            try:
                settlements = extract_settlements(tmp_pdf_path)

                output_name = uploaded_file.name.replace(".pdf", "_QB.xlsx")
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_xlsx:
                    tmp_xlsx_path = tmp_xlsx.name

                write_excel(settlements, tmp_xlsx_path)

                with open(tmp_xlsx_path, "rb") as f:
                    excel_bytes = f.read()

            finally:
                os.unlink(tmp_pdf_path)
                if tmp_xlsx_path and os.path.exists(tmp_xlsx_path):
                    os.unlink(tmp_xlsx_path)

        # ── Summary ───────────────────────────────────────────────────────────
        st.divider()
        st.success("✅ Processing complete!")

        total_rows = sum(len(settlement_to_rows(s)) for s in settlements)
        total_unmapped = sum(
            len(s['unmapped_other_pay']) + len(s['unmapped_reserves'])
            for s in settlements
        )

        unmapped_drivers = [
            s for s in settlements
            if (s['truck_note'] is not None and not DRIVERS.get(s['driver_code'], ('','','','',''))[3]) or
               (s['maint_fund'] is not None and not DRIVERS.get(s['driver_code'], ('','','','',''))[4])
        ]

        col1, col2, col3 = st.columns(3)
        col1.metric("Drivers Processed", len(settlements))
        col2.metric("QB Rows Written", total_rows)
        col3.metric("Unmapped Lines", total_unmapped)

        if unmapped_drivers:
            st.warning(
                "⚠️ **The following drivers have unmapped accounts:**\n" +
                "\n".join(f"- {s['driver_name']}" for s in unmapped_drivers)
            )

        st.divider()

        # ── Driver breakdown ──────────────────────────────────────────────────
        with st.expander("📋 Driver breakdown", expanded=False):
            for s in settlements:
                inv = s['invoice_number'] or 'NO INV#'
                unmapped_count = len(s['unmapped_other_pay']) + len(s['unmapped_reserves'])
                driver_info = DRIVERS.get(s['driver_code'], ('','','','',''))
                flag = " ⚠️" if (
                    (s['truck_note'] is not None and not driver_info[3]) or
                    (s['maint_fund'] is not None and not driver_info[4])
                ) else ""
                st.markdown(
                    f"**{s['driver_name']}**{flag} &nbsp;|&nbsp; "
                    f"Class: `{s['driver_class']}` &nbsp;|&nbsp; "
                    f"INV: `{inv}` &nbsp;|&nbsp; "
                    f"Gross: `${s['gross_pay']:,.2f}` &nbsp;|&nbsp; "
                    f"Unmapped: `{unmapped_count}`"
                )

        # ── Download ──────────────────────────────────────────────────────────
        st.download_button(
            label="⬇️ Download Excel File",
            data=excel_bytes,
            file_name=output_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
        )