import streamlit as st
import pandas as pd
from io import BytesIO

st.title("🔄 BOM Comparison & Supplier Transfer")

# Upload old and new BOM files
old_file = st.file_uploader("Upload OLD BOM (Excel)", type=["xlsx"])
new_file = st.file_uploader("Upload NEW BOM (Excel)", type=["xlsx"])

if old_file and new_file:
    # Load both BOMs
    old_df = pd.read_excel(old_file)
    new_df = pd.read_excel(new_file)

    # Normalize column names (strip spaces, lower)
    old_df.columns = old_df.columns.str.strip()
    new_df.columns = new_df.columns.str.strip()

    # Ensure key exists
    if "MPN" not in old_df.columns or "MPN" not in new_df.columns:
        st.error("❌ Both files must contain an 'MPN' column.")
    else:
        # Columns to copy from old BOM
        transfer_cols = [
            "Supplier", "PO number", "Po qty", "Supplier part number", "Price",
            "Extended price", "Remarks", "ETA", "Currency", "Lead time", "Availability",
            "BCD", "unit price with BCD", "unit price in INR", "Extended price in INR"
        ]

        # Add missing columns in old_df
        for col in transfer_cols:
            if col not in old_df.columns:
                old_df[col] = None

        # Merge on MPN
        merged = pd.merge(
            new_df,
            old_df[["MPN"] + transfer_cols],
            on="MPN",
            how="left"
        )

        # Fill missing supplier info
        merged["Supplier"] = merged["Supplier"].fillna("New Part")

        # Reorder: Insert after Manufacturer
        if "Manufacturer" in merged.columns:
            manufacturer_index = merged.columns.get_loc("Manufacturer")
            before = merged.columns[:manufacturer_index + 1].tolist()
            supplier_block = transfer_cols
            after = [col for col in merged.columns if col not in before + supplier_block]
            merged = merged[before + supplier_block + after]

        # Downloadable output
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            merged.to_excel(writer, index=False, sheet_name="Updated_BOM")

        st.success("✅ BOM processed successfully!")
        st.download_button(
            label="📥 Download Updated BOM",
            data=output.getvalue(),
            file_name="Updated_BOM.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.dataframe(merged.head())
