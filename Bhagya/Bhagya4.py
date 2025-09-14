import streamlit as st
import pandas as pd
from io import BytesIO

st.title("🔄 BOM Supplier Mapping Tool")

# Upload files
old_file = st.file_uploader("Upload OLD BOM", type=["xlsx"])
new_file = st.file_uploader("Upload NEW BOM", type=["xlsx"])

if old_file and new_file:
    # Load both BOMs
    old_df = pd.read_excel(old_file)
    new_df = pd.read_excel(new_file)

    # Normalize headers
    old_df.columns = old_df.columns.str.strip()
    new_df.columns = new_df.columns.str.strip()

    # ✅ Columns to transfer from OLD BOM
    transfer_cols = [
        'Supplier', 'PO number', 'Po qty', 'Supplier part number',
        'Price', 'Extended price', 'Remarks', 'ETA', 'Currency',
        'Lead time', 'Availability', 'BCD',
        'unit price with BCD', 'unit price in INR', 'Extended price in INR'
    ]

    # Check missing columns
    missing = [c for c in transfer_cols if c not in old_df.columns]
    if missing:
        st.error(f"❌ Missing columns in OLD BOM: {missing}")
    else:
        # Extract supplier block
        supplier_df = old_df[['MPN'] + transfer_cols]

        # Merge with NEW BOM on MPN
        merged = pd.merge(
            new_df,
            supplier_df,
            how="left",
            on="MPN",
            suffixes=("", "_old")
        )

        # Handle Remarks:
        # if already in NEW BOM → keep
        # else if missing supplier → "New Part"
        merged['Remarks'] = merged['Remarks'].combine_first(merged['Remarks_old'])
        merged['Remarks'] = merged['Remarks'].fillna("New Part")
        merged.drop(columns=['Remarks_old'], inplace=True, errors="ignore")

        # ✅ Arrange columns: insert supplier block between MPN and Manufacturer
        if "MPN" in new_df.columns and "Manufacturer" in new_df.columns:
            base_cols = list(new_df.columns)
            mpn_idx = base_cols.index("MPN")
            manuf_idx = base_cols.index("Manufacturer")

            new_order = (
                base_cols[:mpn_idx+1] +
                transfer_cols +
                base_cols[manuf_idx:]
            )
            # Deduplicate just in case
            new_order = list(dict.fromkeys(new_order))

            final_df = merged[new_order]
        else:
            st.error("❌ 'MPN' or 'Manufacturer' column missing in NEW BOM")
            final_df = merged

        st.success("✅ Supplier mapping completed successfully!")

        # Preview
        st.subheader("📋 Preview of Mapped BOM")
        st.dataframe(final_df.head(20))

        # Download
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            final_df.to_excel(writer, index=False, sheet_name="MappedBOM")
        st.download_button(
            "📥 Download Mapped BOM",
            data=output.getvalue(),
            file_name="Mapped_BOM.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
