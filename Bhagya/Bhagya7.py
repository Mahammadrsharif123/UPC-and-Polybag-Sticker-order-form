import streamlit as st
import pandas as pd
from io import BytesIO

st.title("🔄 BOM Comparison & Supplier Transfer")

# Upload old and new BOM files
old_file = st.file_uploader("Upload OLD BOM (Excel)", type=["xlsx"])
new_file = st.file_uploader("Upload NEW BOM (Excel)", type=["xlsx"])

def find_mpn_column(df):
    """Find the MPN column in a dataframe (case-insensitive, alias check)."""
    possible_names = ["mpn", "manufacturer part number", "part number"]
    for col in df.columns:
        if col.strip().lower() in possible_names:
            return col
    return None

if old_file and new_file:
    # Load both BOMs
    old_df = pd.read_excel(old_file)
    new_df = pd.read_excel(new_file)

    # Normalize column names (strip spaces)
    old_df.columns = old_df.columns.str.strip()
    new_df.columns = new_df.columns.str.strip()

    # Detect actual MPN column name
    old_mpn_col = find_mpn_column(old_df)
    new_mpn_col = find_mpn_column(new_df)

    if not old_mpn_col or not new_mpn_col:
        st.error("❌ Could not detect 'MPN' column in one of the files. Please check headers.")
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
            old_df[[old_mpn_col] + transfer_cols],
            left_on=new_mpn_col,
            right_on=old_mpn_col,
            how="left"
        )

        # Drop duplicate MPN column after merge
        if old_mpn_col != new_mpn_col and old_mpn_col in merged.columns:
            merged.drop(columns=[old_mpn_col], inplace=True)

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
