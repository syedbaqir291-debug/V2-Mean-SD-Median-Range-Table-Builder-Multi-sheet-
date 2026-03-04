# app_oncology_flexible_mean_sd.py

import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Oncology Mean (SD) Table Builder", layout="wide")

st.title("Oncology Mean (SD) / Median (Range) Table Builder")

st.markdown("""
Upload Excel file (multiple sheets supported).

Steps:
1️⃣ Select OUTER sheets (Mean / Median)  
2️⃣ Select INNER sheets (SD / Range)  
3️⃣ Enter header row number  
4️⃣ Enter start & end column letters  

Categories will be automatically standardized into oncology format.
""")

# ===============================
# FILE UPLOAD
# ===============================
uploaded_file = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])

if not uploaded_file:
    st.stop()

xls = pd.ExcelFile(uploaded_file)
sheets = xls.sheet_names
st.write("Detected Sheets:", sheets)

# ===============================
# SHEET SELECTION
# ===============================
mean_sheets = st.multiselect("Select OUTER value sheet(s)", sheets)
sd_sheets = st.multiselect("Select INNER value sheet(s)", sheets)

if not mean_sheets or not sd_sheets:
    st.warning("Select at least one OUTER and one INNER sheet.")
    st.stop()

# ===============================
# LAYOUT SETTINGS
# ===============================
st.subheader("Layout Settings")

header_row = st.number_input(
    "Header row number (Excel row index starting from 1)",
    min_value=1,
    value=1
)

col1, col2 = st.columns(2)
with col1:
    start_col = st.text_input("Start column letter", value="A")
with col2:
    end_col = st.text_input("End column letter", value="F")

decimals = st.selectbox("Decimal places", [0,1,2], index=0)

# ===============================
# HELPERS
# ===============================
def excel_col_to_index(col_letter):
    num = 0
    for c in col_letter.upper():
        num = num * 26 + (ord(c) - ord('A') + 1)
    return num - 1

def read_multi_sheets(file, sheet_list):
    dfs = []
    start_idx = excel_col_to_index(start_col)
    end_idx = excel_col_to_index(end_col)

    for sheet in sheet_list:
        df = pd.read_excel(file, sheet_name=sheet, header=header_row-1)
        df = df.iloc[:, start_idx:end_idx+1]
        df = df.rename(columns={df.columns[0]: "Category"})
        df = df.set_index("Category")
        df.columns = [str(c).strip() for c in df.columns]
        df = df.apply(pd.to_numeric, errors='coerce')
        dfs.append(df)

    return dfs

# ===============================
# ONCOLOGY STANDARDIZATION
# ===============================
category_order = [
    "Haematological",
    "Gynecological",
    "Urological",
    "Neurological",
    "Breast",
    "Pulmonary",
    "Gastrointestinal",
    "Head & Neck",
    "Thyroid",
    "Sarcoma",
    "Retinoblastoma",
    "Other rare tumors"
]

def standardize_categories(df):
    df = df.copy()
    rename_map = {}

    for cat in df.index:
        cat_clean = str(cat).strip().lower()

        if "non" in cat_clean and "specific" in cat_clean:
            rename_map[cat] = "Other rare tumors"

        elif "hematolog" in cat_clean:
            rename_map[cat] = "Haematological"

        elif "gynecolog" in cat_clean:
            rename_map[cat] = "Gynecological"

        elif "urolog" in cat_clean:
            rename_map[cat] = "Urological"

        elif "neurolog" in cat_clean:
            rename_map[cat] = "Neurological"

        elif "breast" in cat_clean:
            rename_map[cat] = "Breast"

        elif "pulmon" in cat_clean:
            rename_map[cat] = "Pulmonary"

        elif "gastro" in cat_clean:
            rename_map[cat] = "Gastrointestinal"

        elif "head" in cat_clean and "neck" in cat_clean:
            rename_map[cat] = "Head & Neck"

        elif "thyroid" in cat_clean:
            rename_map[cat] = "Thyroid"

        elif "sarcoma" in cat_clean:
            rename_map[cat] = "Sarcoma"

        elif "retino" in cat_clean:
            rename_map[cat] = "Retinoblastoma"

    df = df.rename(index=rename_map)
    return df

# ===============================
# READ + STANDARDIZE
# ===============================
mean_dfs = read_multi_sheets(uploaded_file, mean_sheets)
sd_dfs = read_multi_sheets(uploaded_file, sd_sheets)

mean_dfs = [standardize_categories(df) for df in mean_dfs]
sd_dfs = [standardize_categories(df) for df in sd_dfs]

# ===============================
# COMBINE VALUES
# ===============================
def combine_values(dfs, category, column):
    values = []
    for df in dfs:
        if category in df.index and column in df.columns:
            val = df.loc[category, column]
            if pd.notna(val):
                values.append(val)
            else:
                values.append("")
        else:
            values.append("")
    return values

base_df = mean_dfs[0]
columns = base_df.columns.tolist()

final_df = pd.DataFrame(index=category_order, columns=columns)

for cat in category_order:
    for col in columns:

        mean_vals = combine_values(mean_dfs, cat, col)
        sd_vals = combine_values(sd_dfs, cat, col)

        mean_fmt = [
            f"{float(v):.{decimals}f}" if v != "" else ""
            for v in mean_vals
        ]

        sd_fmt = [
            f"{float(v):.{decimals}f}" if v != "" else ""
            for v in sd_vals
        ]

        mean_str = "-".join(mean_fmt)
        sd_str = "-".join(sd_fmt)

        if mean_str == "" and sd_str == "":
            final_df.loc[cat, col] = "–"
        elif mean_str != "" and sd_str == "":
            final_df.loc[cat, col] = mean_str
        elif mean_str == "" and sd_str != "":
            final_df.loc[cat, col] = f"({sd_str})"
        else:
            final_df.loc[cat, col] = f"{mean_str} ({sd_str})"

# ===============================
# DISPLAY
# ===============================
st.subheader("Final Oncology Standardized Table")
st.dataframe(final_df.reset_index().rename(columns={"index":"Category"}), use_container_width=True)

# ===============================
# EXPORT
# ===============================
def to_excel_bytes(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Mean_SD_Table", index=False)
    return output.getvalue()

excel_bytes = to_excel_bytes(
    final_df.reset_index().rename(columns={"index":"Category"})
)

st.download_button(
    label="Download Excel",
    data=excel_bytes,
    file_name="Oncology_Mean_SD_Table.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.success("Oncology standardized table generated successfully ✅")
