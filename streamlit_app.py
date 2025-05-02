
import streamlit as st
import os
import pandas as pd
import xlsxwriter
from collections import defaultdict
from io import BytesIO
import zipfile

st.set_page_config(page_title="Metrics Comparator", layout="wide")

st.title("📊 Metrics Comparator")

st.markdown("Upload **Original** and **Reduced TM** metrics Excel files below. Files will be paired automatically by name.")

original_files = st.file_uploader("Original Project Files", type="xlsx", accept_multiple_files=True, key="original")
reduced_files = st.file_uploader("Reduced TM Project Files", type="xlsx", accept_multiple_files=True, key="reduced")

def base_name(name):
    return name.replace("_Reduced TM", "").replace("_Metrics.xlsx", "").replace(".xlsx", "").strip()

def extract_header_and_table(df):
    start_row = None
    for idx, row in df.iterrows():
        if str(row.iloc[0]).strip().lower() == "total count":
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
    return table

def compare_files(original_files, reduced_files):
    paired = {}
    for file in original_files:
        name = base_name(file.name)
        paired[name] = {'original': file, 'reduced': None}
    for file in reduced_files:
        name = base_name(file.name)
        if name in paired:
            paired[name]['reduced'] = file

    buffer = BytesIO()
    workbook = xlsxwriter.Workbook(buffer)
    worksheet = workbook.add_worksheet()
    row_cursor = 0

    header_format = workbook.add_format({'bold': True, 'font_color': 'black'})
    center_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
    left_format = workbook.add_format({'align': 'left', 'valign': 'vcenter', 'border': 1})
    bold_format = workbook.add_format({'bold': True, 'border': 1})
    red_bold = workbook.add_format({'bold': True, 'font_color': 'red'})
    regular = workbook.add_format({'font_color': 'black'})
    header_bg = workbook.add_format({'bold': True, 'bg_color': '#D9D9D9', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True})

    delta_summary = {}

    for name, files in paired.items():
        if files['original'] and files['reduced']:
            df1 = pd.read_excel(files['original'], header=None)
            df2 = pd.read_excel(files['reduced'], header=None)
            t1 = extract_header_and_table(df1)
            t2 = extract_header_and_table(df2)

            worksheet.merge_range(row_cursor, 0, row_cursor, 6, "", center_format)
            worksheet.merge_range(row_cursor, 8, row_cursor, 14, "", center_format)
            worksheet.write_rich_string(row_cursor, 0, red_bold, "File Name 1: ", regular, files['original'].name, red_bold, " Original Project", center_format)
            worksheet.write_rich_string(row_cursor, 8, red_bold, "File Name 2: ", regular, files['reduced'].name, red_bold, " New Project", center_format)
            row_cursor += 1

            headers = ["Initial", "Segments", "Words", "Word %", "Characters", "Characters %", "Characters excluding spaces"]
            for idx, h in enumerate(headers):
                worksheet.write(row_cursor, idx, h, header_bg)
                worksheet.write(row_cursor, idx+8, h, header_bg)
                worksheet.write(row_cursor, idx+16, h + " Δ" if idx > 0 else "Metric", header_bg)
            row_cursor += 1

            merged = pd.merge(t1, t2, on="Initial", how="inner")
            for _, row in merged.iterrows():
                metric = row['Initial']
                for j, col in enumerate(headers):
                    val1 = row.get(col + "_x", row.get(col, ""))
                    val2 = row.get(col + "_y", row.get(col, ""))
                    val1_fmt = bold_format if metric in ["Translation memory matching", "Internal matching"] else (left_format if j == 0 else center_format)
                    val2_fmt = val1_fmt

                    worksheet.write(row_cursor, j, "" if pd.isna(val1) else val1, val1_fmt)
                    worksheet.write(row_cursor, j+8, "" if pd.isna(val2) else val2, val2_fmt)

                    if j == 0:
                        worksheet.write(row_cursor, j+16, metric, val1_fmt)
                    else:
                        try:
                            if val1 == val2:
                                diff = float(val1)
                                fmt = workbook.add_format({"align": "center", "valign": "vcenter", "border": 1, "font_color": "black"})
                                worksheet.write(row_cursor, j+16, str(int(diff)), fmt)
                            else:
                                diff = float(val2) - float(val1)
                                fmt = workbook.add_format({"align": "center", "valign": "vcenter", "border": 1, "font_color": "green" if diff > 0 else "red"})
                                worksheet.write(row_cursor, j+16, f"+{int(diff)}" if diff > 0 else str(int(diff)), fmt)
                            if metric not in delta_summary:
                                delta_summary[metric] = defaultdict(float)
                            delta_summary[metric][j] += diff
                        except:
                            worksheet.write(row_cursor, j+16, "0", center_format)
                row_cursor += 1
            row_cursor += 1

    # Combined summary table
    if delta_summary:
        worksheet.write(row_cursor, 0, "📊 Combined Delta Summary", red_bold)
        row_cursor += 1
        summary_headers = ["Metric", "Segments Δ", "Words Δ", "Word % Δ", "Characters Δ", "Characters % Δ", "Characters excluding spaces Δ"]
        for j, h in enumerate(summary_headers):
            worksheet.write(row_cursor, j, h, header_bg)
        worksheet.set_row(row_cursor, 30)
        worksheet.set_column(6, 6, 10)
        worksheet.set_column(7, 7, 3)
        worksheet.set_column(14, 14, 10)
        worksheet.set_column(15, 15, 3)
        worksheet.set_column(22, 22, 10)
        row_cursor += 1
        for metric in delta_summary:
            metric_lc = metric.strip().lower()
            is_special = metric_lc in ["total count", "non-translatable", "no matching"]
            row_format = workbook.add_format({"align": "center", "valign": "vcenter", "border": 1, "font_color": "black"}) if is_special else left_format
            worksheet.write(row_cursor, 0, metric, row_format)
            for j in range(1, len(summary_headers)):
                total = delta_summary[metric].get(j, 0)
                if is_special:
                    fmt = workbook.add_format({"align": "center", "valign": "vcenter", "border": 1, "font_color": "black"})
                    worksheet.write(row_cursor, j, str(int(total)), fmt)
                else:
                    fmt = workbook.add_format({"align": "center", "valign": "vcenter", "border": 1, "font_color": "green" if total > 0 else "red" if total < 0 else "black"})
                    worksheet.write(row_cursor, j, f"+{int(total)}" if total > 0 else str(int(total)), fmt)
            row_cursor += 1

    workbook.close()
    buffer.seek(0)
    return buffer

if st.button("🧠 Generate Report"):
    if original_files and reduced_files:
        output = compare_files(original_files, reduced_files)
        st.success("✅ Report generated!")
        st.download_button(label="📥 Download Excel Report", data=output, file_name="metrics_comparison.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.error("Please upload both original and reduced TM files.")
