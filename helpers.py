import streamlit as st
import matplotlib.pyplot as plt
import pandas as pd
from PIL import Image
from openpyxl.drawing.image import Image as XLImage
import os
import xlwings as ws
import re
import numpy as np

# Hàm kiểm tra value có phải là năm hợp lệ không
def is_year(x):
    if isinstance(x, float) and x.is_integer():
        x = int(x)
    x = str(x).strip()
    if len(x) == 4 and x.isdigit():
        y=int(x)
        return 1900 <= y <= 2100
    return False

# Hàm tạo và định dạng lại header theo hàng chứa năm (năm trong file data không trùng định dạng (vừa int, vừa float, vừa string) nên cần phải chuyển thành string hết)
def year_header(data):
    header = []
    for y in data.iloc[0]:
        if is_year(y):
            header.append(str(int(y))) 
        else:
            header.append(str(y))
    return header

def select_year_range(data):
    data_years = []
    for y in data.columns:
        if is_year(y):
            data_years.append(str(int(y)))

    begin_year = st.selectbox("Chọn năm bắt đầu:", data_years)
    end_options = [e_y for e_y in data_years if e_y >= begin_year]# Danh sách cho năm kết thúc

        # Tạo session_state để lưu giá trị năm kết thúc
    if "end_year" not in st.session_state:
       st.session_state.end_year = end_options[0]

        # Thay end_year mới khi list end_options thay đổi do người dùng thay đổi end_year 
    if st.session_state.end_year not in end_options:
        st.session_state.end_year = end_options[0]

    def_end_year_index = end_options.index(st.session_state.end_year) # Gán index mặc định cho giá trị end_year theo giá trị người dùng nhập trong lần rerun trước
    end_year = st.selectbox("Chọn năm kết thúc:", end_options, index = def_end_year_index)

    st.session_state.end_year = end_year # Cập nhật session state

    selected_years = [s_y for s_y in data_years if begin_year <= s_y <= end_year] # Năm hiển thị

    return selected_years

def plot_chart(label, folder_path, chart_type, x_col, y_col):
    if chart_type == "Line Chart":
        fig, ax = plt.subplots(figsize=(10, 5))
        ax.plot(x_col, y_col, marker='o', linewidth=2)
        ax.set_title(f"Biểu đồ {label}", fontsize=14)
        ax.set_xlabel("Năm", fontsize=12)
        ax.set_ylabel("Giá trị", fontsize=12)
        ax.grid(True)

        chart_name = f"{chart_type}_of_{label.strip().replace(" ", "_")}.png"
        fig.savefig(f"./{folder_path}/{chart_name}")

    elif chart_type == "Bar Chart":
        fig, ax = plt.subplots(figsize=(10, 5))
        ax.bar(x_col, y_col)
        ax.set_title(f"Biểu đồ {label}", fontsize=14)
        ax.set_xlabel("Năm", fontsize=12)
        ax.set_ylabel("Giá trị", fontsize=12)
        ax.grid(True)

        chart_name = f"{chart_type}_of_{label.strip().replace(" ", "_")}.png"
        fig.savefig(f"./{folder_path}/{chart_name}")

    return f"F:/CSDATA10Cybersoft/do_an_cuoi_khoa/{folder_path}/{chart_name}"

def safe_div(a, b):
    return np.where(b == 0, np.nan, a / b)

def render_chart_from_session_state(session_state_report, chart_folder_path):
    for i, report in enumerate(session_state_report):
        chart_path = report["chart_path"]
        if chart_path.endswith(('.png', 'jpg', 'jpeg')):
            file_path = os.path.join(chart_folder_path, chart_path)

            image = Image.open(file_path)
            st.image(image, caption= chart_path, use_column_width=True)

def calculate_financial_ratios(data, selected_years):
    ratio_df = pd.DataFrame(index = selected_years)

    ratio_df["Current Ratio"] = safe_div(
        data["tài sản ngắn hạn"],
        data["nợ ngắn hạn"]
    )

    ratio_df["Quick Ratio"] = safe_div(
        data["tài sản ngắn hạn"] - data["hàng tồn kho"],
        data["nợ ngắn hạn"]
    )

    ratio_df["Cash Ratio"] = safe_div(
        data["tiền và tương đương tiền"],
        data["nợ ngắn hạn"]
    )

    # -------- Cơ cấu tài chính --------
    ratio_df["Debt to Total Assets"] = safe_div(
        data["nợ phải trả"],
        data["tổng tài sản"]
    )

    ratio_df["Debt to Equity"] = safe_div(
        data["nợ phải trả"],
        data["vốn chủ sở hữu"]
    )

    ratio_df["Equity Ratio"] = safe_div(
        data["vốn chủ sở hữu"],
        data["tổng tài sản"]
    )

    ratio_df["Financial Leverage"] = safe_div(
        data["tổng tài sản"],
        data["vốn chủ sở hữu"]
    )

    # -------- Chất lượng tài sản --------
    ratio_df["Tỷ trọng TSNH"] = safe_div(
        data["tài sản ngắn hạn"],
        data["tổng tài sản"]
    )

    ratio_df["Tỷ trọng TSCĐ"] = safe_div(
        data["tài sản cố định"],
        data["tổng tài sản"]
    )

    return ratio_df


def generate_excel_report(data, reports, report_name):
    """
    Generates an Excel report containing:
    - Original dataset
    - Pivot tables from aggregation
    - Charts (saved as images)
    - AI-generated insights

    Args:
    - data (pd.DataFrame): The original dataset.
    - reports (list): A list of reports with pivot tables, charts, and insights.
    - report_name (str): The output Excel file name (without extension).

    Returns:
    - str: The absolute path to the generated report.
    """

    report_filename = f"{report_name}.xlsx"
    report_path = os.path.abspath(report_filename)

    # Kiểm tra file đã tồn tại chưa
    file_exists = os.path.exists(report_path)

    if file_exists:
        writer = pd.ExcelWriter(report_path, engine='openpyxl', mode='a', if_sheet_exists="replace")
    else:
        writer = pd.ExcelWriter(report_path, engine='openpyxl', mode='w')

    with writer:
        if not file_exists:
            data.to_excel(writer, sheet_name="Datasource", index=False)

        for report in reports:
            sheet_name = report["sheet_name"]

            # Save the pivot table to Excel
            report["pivot_table"].to_excel(writer, sheet_name=sheet_name, index=False)

            # Get the worksheet to add images and insights
            worksheet = writer.sheets[sheet_name]

            # Add Chart Image if Exists
            chart_path = report["chart_path"]
            if os.path.exists(chart_path):
                img = XLImage(chart_path)
                img.width = img.width * 0.7
                img.height = img.height * 0.7
                worksheet.add_image(img, "D7")

def remove_report_file(report_name):
    report_filename = f"{report_name}.xlsx"
    if os.path.exists(report_filename):
        os.remove(report_filename)
