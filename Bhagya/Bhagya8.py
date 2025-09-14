import streamlit as st
import pandas as pd
from io import BytesIO

st.title("📑 BOM Comparator & Supplier Details Filler")

# --- Function to detect MPN column safely ---
def find_mpn_column(df):
    """Find the MPN column in a dataframe (case-insensitive, alias check)."""
    possible_names = ["mpn", "manufacturer part number", "part number"]
    for col in df.columns:
        col_str = str(col).strip().lower()   # Convert to string
        if col_str in possible_names:
            return col
    return None

# --- Upload old BOM ---
old_file = st.file_uploader("Upload OLD BOM", type=["xlsx"])
# --- Upload new BOM ---
new_file = st.file_uploader("Upload NEW BOM", type=["xlsx"])

if old_file and new_file:
    # Read both Excel files
    old_df = pd.read_excel(old_file)
    new_df = pd.read_excel(new_file)

    # Normalize column names
    old_df.columns = [str(c).strip() for c in old_df.columns]
    new_df.columns = [str(c).strip() for c in new_df.columns]

    # Detect MPN columns
    old_mpn_col = find_mpn_column(old_df)
    new_mpn_col = find_mpn_column(new_df)

    if not old_mpn_col or not new_mpn_col:
        st.error("❌ Both files must contain an 'MPN' (or similar) column.")
        st.stop()

    st.success(f"✅ Detected MPN column → OLD: {old_mpn_col}, NEW: {new_mpn_col}")

    # Define columns to transfer from old BOM
    transfer_cols = [
        "Supplier", "PO number", "Po qty", "Supplier part number",
        "Price", "Extended price", "Remarks", "ETA", "Currency",
        "Lead time", "Availability", "BCD", "unit price with BCD",
        "unit price in INR", "Extended price in INR"
    ]

    # Keep only required columns from old BOM
    old_subset = old_df[[old_mpn_col] + [c for c in transfer_cols if c in old_df.columns]]

    # Merge on MPN
    merged = pd.merge(
        new_df,
        old_subset,
        how="left",
        left_on=new_mpn_col,
        right_on=old_mpn_col,
        suffixes=("", "_old")
    )

    # Drop duplicate merge column
    if old_mpn_col != new_mpn_col:
        merged = merged.drop(columns=[old_mpn_col])

    # Insert supplier-related columns right after Manufacturer
    if "Manufacturer" in merged.columns:
        manuf_index = merged.columns.get_loc("Manufacturer") + 1
    else:
        manuf_index = merged.columns.get_loc(new_mpn_col) + 1

    for i, col in enumerate(transfer_cols):
        if col in merged.columns:
            # Move column to desired position
            cols = list(merged.columns)
            cols.insert(manuf_index + i, cols.pop(cols.index(col)))
            merged = merged[cols]

    # Mark new parts where supplier is missing
    if "Supplier" in merged.columns:
        merged["Remarks"] = merged["Remarks"].fillna("")
        merged.loc[merged["Supplier"].isna(), "Remarks"] = "New Part"

    # --- Preview result ---
    st.subheader("🔍 Preview of Updated BOM")
    st.dataframe(merged)

    # --- Export to Excel ---
    def to_excel(df):
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Updated BOM")
        return output.getvalue()

    st.download_button(
        label="📥 Download Updated BOM",
        data=to_excel(merged),
        file_name="Updated_BOM.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
