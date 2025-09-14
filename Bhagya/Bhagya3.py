import streamlit as st
import pandas as pd
from io import BytesIO

st.title("🔄 BOM Supplier Mapping Tool")

# --- Upload files ---
old_file = st.file_uploader("Upload OLD BOM", type=["xlsx"])
new_file = st.file_uploader("Upload NEW BOM", type=["xlsx"])

if old_file and new_file:
    # Load both BOMs
    old_df = pd.read_excel(old_file)
    new_df = pd.read_excel(new_file)

    st.write("✅ Files uploaded successfully")

    # Normalize column names (strip spaces, lower)
    old_df.columns = old_df.columns.str.strip()
    new_df.columns = new_df.columns.str.strip()

    # Columns from OLD BOM (CD to CR → adjust names if different)
    transfer_cols = [
        'Supplier', 'PO number', 'PO qty', 'Supplier part number',
        'Price', 'Extended price', 'Remarks', 'ETA', 'Currency',
        'Lead time', 'Availability', 'BCD', 'unit price with BCD',
        'unit price in INR', 'Extended price in INR'
    ]

    # Check missing columns
    missing = [c for c in transfer_cols if c not in old_df.columns]
    if missing:
        st.error(f"❌ Missing columns in OLD BOM: {missing}")
    else:
        # Select required cols from OLD
        supplier_df = old_df[['Manufacturer Part Number'] + transfer_cols]

        # Merge on MPN
        merged = pd.merge(
            new_df,
            supplier_df,
            how='left',
            left_on='Manufacturer Part Number',
            right_on='Manufacturer Part Number'
        )

        # Add remarks for new parts
        merged['Remarks'] = merged['Remarks'].fillna("New Part")

        # Reorder: put transfer cols right after MPN
        cols = list(new_df.columns)
        if "Manufacturer Part Number" in cols and "Manufacturer" in cols:
            mpn_idx = cols.index("Manufacturer Part Number")
            manuf_idx = cols.index("Manufacturer")

            # Build new order
            new_order = (
                cols[:mpn_idx+1] +
                transfer_cols +
                cols[manuf_idx:]
            )

            # Drop duplicates (since we appended)
            new_order = list(dict.fromkeys(new_order))

            final_df = merged[new_order]
        else:
            st.error("❌ 'Manufacturer Part Number' or 'Manufacturer' column missing in NEW BOM")
            final_df = merged

        st.success("✅ Mapping completed")

        # --- Download ---
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            final_df.to_excel(writer, index=False, sheet_name="MappedBOM")
        st.download_button(
            label="📥 Download Mapped BOM",
            data=output.getvalue(),
            file_name="Mapped_BOM.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
