import streamlit as st
import pandas as pd
from io import BytesIO
import difflib

st.title("📑 BOM Comparator & Supplier Filler")

# --- Detect MPN column dynamically ---
def find_mpn_column(df):
    candidates = ["mpn", "manufacturer part number", "part number", "mfr p/n", "mfr part number"]
    df_cols = [str(c).strip().lower() for c in df.columns]

    for cand in candidates:
        if cand in df_cols:
            return df.columns[df_cols.index(cand)]

    close = difflib.get_close_matches("mpn", df_cols, n=1, cutoff=0.5)
    if close:
        return df.columns[df_cols.index(close[0])]

    return None

# --- Upload files ---
old_file = st.file_uploader("Upload OLD BOM", type=["xlsx"])
new_file = st.file_uploader("Upload NEW BOM", type=["xlsx"])

if old_file and new_file:
    old_df = pd.read_excel(old_file)
    new_df = pd.read_excel(new_file)

    old_df.columns = [str(c).strip() for c in old_df.columns]
    new_df.columns = [str(c).strip() for c in new_df.columns]

    # Detect MPN columns
    old_mpn_col = find_mpn_column(old_df)
    new_mpn_col = find_mpn_column(new_df)

    if not old_mpn_col or not new_mpn_col:
        st.error("❌ Could not detect MPN column. Please check your BOM headers.")
        st.write("OLD BOM Columns:", list(old_df.columns))
        st.write("NEW BOM Columns:", list(new_df.columns))
        st.stop()

    # Columns to transfer
    transfer_cols = [
        "Supplier", "PO number", "PO qty", "Supplier part number",
        "Price", "Extended price", "Remarks", "ETA", "Currency",
        "Lead time", "Availability", "BCD", "unit price with BCD",
        "unit price in INR", "Extended price in INR"
    ]

    # Subset old BOM with only relevant columns
    available_transfer = [c for c in transfer_cols if c in old_df.columns]
    old_subset = old_df[[old_mpn_col] + available_transfer]

    # Merge on MPN
    merged = pd.merge(
        new_df,
        old_subset,
        how="left",
        left_on=new_mpn_col,
        right_on=old_mpn_col,
        suffixes=("", "_old")
    )

    if old_mpn_col != new_mpn_col:
        merged = merged.drop(columns=[old_mpn_col])

    # Insert after Manufacturer
    if "Manufacturer" in merged.columns:
        insert_at = merged.columns.get_loc("Manufacturer") + 1
    else:
        insert_at = merged.columns.get_loc(new_mpn_col) + 1

    for i, col in enumerate(available_transfer):
        if col in merged.columns:
            cols = list(merged.columns)
            cols.insert(insert_at + i, cols.pop(cols.index(col)))
            merged = merged[cols]

    # Mark new parts
    if "Supplier" in merged.columns:
        merged["Remarks"] = merged["Remarks"].fillna("")
        merged.loc[merged["Supplier"].isna(), "Remarks"] = "New Part"

    # --- Preview ---
    st.subheader("🔍 Preview of Updated BOM")
    st.dataframe(merged)

    # --- Download ---
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
