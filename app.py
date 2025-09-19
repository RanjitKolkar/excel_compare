import streamlit as st
import pandas as pd
import re
from datetime import timedelta
from fuzzywuzzy import fuzz
from io import BytesIO

st.set_page_config(page_title="Excel Comparator", page_icon="📊", layout="wide")

st.title("📊 Generalized Excel Comparator")

st.markdown("""
This tool lets you **compare two Excel sheets** based on selected columns.  
It automatically handles:
- 🗓 Date formats (`02-09-2025`, `9/2/2025 22:41:26`, etc.)  
- 🔢 Numeric differences (`500`, `500.00`, `500,00`)  
- 📝 Text differences (fuzzy matching option)  
""")

# --- Helper functions ---
def clean_date(val):
    """Normalize dates with multiple formats"""
    if pd.isna(val):
        return None
    try:
        return pd.to_datetime(str(val), errors="coerce").date()
    except:
        return None

def clean_number(val):
    """Normalize numbers: remove commas, convert to float"""
    if pd.isna(val):
        return None
    try:
        return float(str(val).replace(",", "").strip())
    except:
        return None

def compare_values(v1, v2, col_type, fuzzy=False):
    """Compare values by type"""
    if pd.isna(v1) or pd.isna(v2):
        return False

    if col_type == "Date":
        return v1 == v2 or v1 == v2 + timedelta(days=1) or v1 == v2 - timedelta(days=1)

    if col_type == "Number":
        return abs(v1 - v2) < 1e-6  # tolerance

    if col_type == "Text":
        if fuzzy:
            return fuzz.partial_ratio(str(v1).lower(), str(v2).lower()) >= 80
        else:
            return str(v1).strip().lower() == str(v2).strip().lower()

    return str(v1) == str(v2)

# --- File uploads ---
col1, col2 = st.columns(2)
with col1:
    file1 = st.file_uploader("📂 Upload Excel 1", type=["xlsx"])
with col2:
    file2 = st.file_uploader("📂 Upload Excel 2", type=["xlsx"])

if file1 and file2:
    df1 = pd.read_excel(file1)
    df2 = pd.read_excel(file2)

    st.subheader("🔎 Preview of Data")
    c1, c2 = st.columns(2)
    with c1:
        st.write("**Excel 1**")
        st.dataframe(df1.head())
    with c2:
        st.write("**Excel 2**")
        st.dataframe(df2.head())

    st.markdown("---")

    # Column matching selection
    st.subheader("⚙️ Select Columns to Compare")

    col_map = {}
    for col in df1.columns:
        match = st.selectbox(
            f"🔗 Match column from Excel 1: `{col}`",
            ["-- None --"] + list(df2.columns),
            index=0,
        )
        if match != "-- None --":
            col_type = st.radio(
                f"Type for comparison `{col}` vs `{match}`",
                ["Text", "Number", "Date"],
                horizontal=True,
                key=col,
            )
            fuzzy = False
            if col_type == "Text":
                fuzzy = st.checkbox(f"Enable fuzzy match for `{col}`", key=f"{col}_fuzzy")
            col_map[col] = (match, col_type, fuzzy)

    if st.button("🚀 Run Comparison"):
        results = []

        for i, row1 in df1.iterrows():
            match_found = False
            best_match = None

            for j, row2 in df2.iterrows():
                comparisons = []
                for col1, (col2, col_type, fuzzy) in col_map.items():
                    v1, v2 = row1[col1], row2[col2]

                    # Normalize
                    if col_type == "Date":
                        v1, v2 = clean_date(v1), clean_date(v2)
                    elif col_type == "Number":
                        v1, v2 = clean_number(v1), clean_number(v2)

                    comparisons.append(compare_values(v1, v2, col_type, fuzzy))

                if all(comparisons):  # all selected columns matched
                    match_found = True
                    best_match = row2
                    break

            results.append({
                "Excel1_Index": i,
                "Excel1_Data": dict(row1),
                "Matched": "✅ Yes" if match_found else "❌ No",
                "Excel2_Match": dict(best_match) if best_match is not None else None
            })

        # Build result dataframe
        flat_results = []
        for r in results:
            row = {"Excel1_Index": r["Excel1_Index"], "Matched": r["Matched"]}
            for k, v in r["Excel1_Data"].items():
                row[f"Excel1_{k}"] = v
            if r["Excel2_Match"]:
                for k, v in r["Excel2_Match"].items():
                    row[f"Excel2_{k}"] = v
            flat_results.append(row)

        result_df = pd.DataFrame(flat_results)

        st.subheader("📋 Comparison Results")
        st.dataframe(result_df)

        # Download results
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            result_df.to_excel(writer, index=False, sheet_name="Results")
        st.download_button("📥 Download Results", data=output.getvalue(),
                           file_name="Excel_Comparison.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
