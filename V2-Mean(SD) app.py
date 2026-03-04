# app_flexible_multi_mean_sd.py

import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Flexible Mean (SD) Table Builder", layout="wide")

st.title("Flexible Mean (SD) / Median(Range) Table Builder")
st.markdown("""
Upload an Excel workbook (multiple sheets supported).

Steps:
1️⃣ Select sheets for OUTER values (Mean / Median)  
2️⃣ Select sheets for INNER values (SD / Range)  
3️⃣ Enter header row number  
4️⃣ Enter start & end column letters  

The app will combine values cell-wise using dash `-`.
""")

# =============================
# File Upload
# =============================
uploaded_file = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])

if not uploaded_file:
    st.info("Please upload an Excel file.")
    st.stop()

xls = pd.ExcelFile(uploaded_file)
sheets = xls.sheet_names

st.write("Detected Sheets:", sheets)

# =============================
# Sheet Selection
# =============================
mean_sheets = st.multiselect("Select OUTER value sheet(s) (Mean / Median)", sheets)
sd_sheets = st.multiselect("Select INNER value sheet(s) (SD / Range)", sheets)

if not mean_sheets or not sd_sheets:
    st.warning("Please select at least one sheet for both OUTER and INNER values.")
    st.stop()

# =============================
# Layout Controls
# =============================
st.subheader("Layout Settings")

header_row = st.number_input(
    "Enter header row number (Excel row number, starts from 1)",
    min_value=1,
    value=1
)

col1, col2 = st.columns(2)

with col1:
    start_col = st.text_input("Enter START column letter (example: A)", value="A")

with col2:
    end_col = st.text_input("Enter END column letter (example: F)", value="F")

decimals = st.selectbox("Select decimal places", [0,1,2], index=0)

# =============================
# Helpers
# =============================
def excel_col_to_index(col_letter):
    num = 0
    for c in col_letter.upper():
        num = num * 26 + (ord(c) - ord('A') + 1)
    return num - 1  # zero-based index

def read_multi_sheets(file, sheet_list, header_row, start_col, end_col):
    dfs = []

    start_idx = excel_col_to_index(start_col)
    end_idx = excel_col_to_index(end_col)

    for sheet in sheet_list:
        df = pd.read_excel(file, sheet_name=sheet, header=header_row-1)

        # Keep selected columns only
        df = df.iloc[:, start_idx:end_idx+1]

        # Rename first column to Category
        df = df.rename(columns={df.columns[0]: "Category"})
        df = df.set_index("Category")

        # Clean column names
        df.columns = [str(c).strip() for c in df.columns]

        # Convert numeric safely
        df = df.apply(pd.to_numeric, errors='coerce')

        dfs.append(df)

    return dfs

# =============================
# Read Sheets
# =============================
mean_dfs = read_multi_sheets(uploaded_file, mean_sheets, header_row, start_col, end_col)
sd_dfs = read_multi_sheets(uploaded_file, sd_sheets, header_row, start_col, end_col)

# =============================
# Combine Values
# =============================
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

# Use first sheet structure as template
base_df = mean_dfs[0]

categories = base_df.index.tolist()
columns = base_df.columns.tolist()

final_df = pd.DataFrame(index=categories, columns=columns)

for cat in categories:
    for col in columns:

        mean_vals = combine_values(mean_dfs, cat, col)
        sd_vals = combine_values(sd_dfs, cat, col)

        # Format mean values
        mean_fmt = []
        for v in mean_vals:
            if v == "":
                mean_fmt.append("")
            else:
                mean_fmt.append(f"{float(v):.{decimals}f}")

        sd_fmt = []
        for v in sd_vals:
            if v == "":
                sd_fmt.append("")
            else:
                sd_fmt.append(f"{float(v):.{decimals}f}")

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

# =============================
# Display
# =============================
st.subheader("Final Combined Table")
st.dataframe(final_df.reset_index().rename(columns={"index":"Category"}), use_container_width=True)

# =============================
# Excel Download
# =============================
def to_excel_bytes(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Mean_SD_Table")
    return output.getvalue()

excel_bytes = to_excel_bytes(final_df.reset_index().rename(columns={"index":"Category"}))

st.download_button(
    label="Download as Excel",
    data=excel_bytes,
    file_name="Mean_SD_Table.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.success("Table Generated Successfully ✅")
