import streamlit as st
import pandas as pd
import xlsxwriter
import io
import math

def extract_header_and_table(df):
    file_name = ""
    target_language = ""
    date_info = ""
    start_row = None
    found_file_row = None

    for idx, row in df.iterrows():
        val = str(row.iloc[0]).strip().lower()
        if val == "file" and found_file_row is None:
            file_name = str(row.iloc[1]).strip()
            found_file_row = idx
        elif val == "target language" and found_file_row is not None and idx > found_file_row:
            target_language = str(row.iloc[1]).strip()
        elif val == "date" and found_file_row is not None and idx > found_file_row:
            date_info = str(row.iloc[1]).strip()
        if str(row.iloc[0]).strip() == "Total count":
            start_row = idx - 1
            break

    if start_row is None:
        raise ValueError("No 'Total count' found.")

    full_table_data = []
    for idx in range(start_row, len(df)):
        row_data = df.iloc[idx, :7].tolist()
        full_table_data.append(row_data)
        if str(df.iloc[idx, 0]).strip() == "No matching":
            break

    table = pd.DataFrame(full_table_data, columns=[
        "Initial", "Segments", "Words", "Word %", "Characters", "Characters %", "Characters excluding spaces"
    ])
    table = table[~table['Initial'].isin(['Initial', 'Metric'])]
    return file_name, target_language, date_info, table

def process_files(original_file, new_file):
    df_original = pd.read_excel(original_file, header=None)
    df_new = pd.read_excel(new_file, header=None)

    file1_name, target1, date1, table1 = extract_header_and_table(df_original)
    file2_name, target2, date2, table2 = extract_header_and_table(df_new)

    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    worksheet = workbook.add_worksheet()

    header_format = workbook.add_format({'bold': True, 'bg_color': '#D9D9D9', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True})
    center_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
    left_format = workbook.add_format({'align': 'left', 'valign': 'vcenter', 'border': 1})
    bold_format = workbook.add_format({'bold': True, 'border': 1})
    red_bold = workbook.add_format({'bold': True, 'font_color': 'red'})
    regular = workbook.add_format({'font_color': 'black'})

    worksheet.merge_range(0, 0, 0, 6, "", center_format)
    worksheet.merge_range(0, 8, 0, 14, "", center_format)

    worksheet.write_rich_string(0, 0,
        red_bold, "File Name 1: ",
        regular, file1_name,
        red_bold, " Original Project",
        center_format)
    worksheet.write_rich_string(0, 8,
        red_bold, "File Name 2: ",
        regular, file2_name,
        red_bold, " New Project",
        center_format)

    headers = ["Initial", "Segments", "Words", "Word %", "Characters", "Characters %", "Characters excluding spaces"]
    for idx, h in enumerate(headers):
        worksheet.write(1, idx, h, header_format)
        worksheet.write(1, idx+8, h, header_format)

    worksheet.set_row(1, 30)

    for i, row in table1.iterrows():
        for j, val in enumerate(row):
            val = "" if pd.isna(val) or (isinstance(val, float) and math.isnan(val)) else val
            fmt = bold_format if j == 0 and row['Initial'] in ["Translation memory matching", "Internal matching"] else (left_format if j == 0 else center_format)
            worksheet.write(2+i, j, val, fmt)

    for i, row in table2.iterrows():
        for j, val in enumerate(row):
            val = "" if pd.isna(val) or (isinstance(val, float) and math.isnan(val)) else val
            fmt = bold_format if j == 0 and row['Initial'] in ["Translation memory matching", "Internal matching"] else (left_format if j == 0 else center_format)
            worksheet.write(2+i, j+8, val, fmt)

    worksheet.set_column(0, 0, 26.71)
    worksheet.set_column(1, 6, 12)
    worksheet.set_column(7, 7, 3)
    worksheet.set_column(8, 8, 26.71)
    worksheet.set_column(9, 14, 12)
    worksheet.set_column(6, 6, 15.57)
    worksheet.set_column(14, 14, 15.57)

    workbook.close()
    output.seek(0)
    return output

st.title("📊 Metrics Comparator Web App")

original_file = st.file_uploader("Upload Original Project Excel (.xlsx)", type="xlsx")
new_file = st.file_uploader("Upload New Project Excel (.xlsx)", type="xlsx")

if st.button("Compare Files") and original_file and new_file:
    with st.spinner('Processing files...'):
        output = process_files(original_file, new_file)
        st.success("✅ Comparison Complete!")
        st.download_button("Download Comparison Summary", data=output, file_name="Comparison_Summary.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")