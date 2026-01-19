import streamlit as st
import pandas as pd
import os
from PIL import Image
from helpers import *

st.title("Phân tích cơ bản")

file=st.sidebar.file_uploader("Upload a file", type=(["xlsx"]))

if "balance_sheet_report" not in st.session_state:
    st.session_state.balance_sheet_report = []

if "ratio_report" not in st.session_state:
    st.session_state.ratio_report = []

if "no" not in st.session_state:
    st.session_state.no = 1

chart_folder_name = "charts"
chart_folder_path = f"./{chart_folder_name}"

if not os.path.exists(chart_folder_name):
    os.makedirs(chart_folder_name)

if file is not None:
    data=pd.read_excel(file)

    st.subheader("Các lựa chọn tổng hợp dữ liệu")
    data.dropna(inplace=True)

    # Chỉnh header của file thành các cột năm
    header_row = year_header(data)
    data.columns = header_row # set header mới
    data = data.iloc[1:].reset_index(drop=True)# xóa header cũ và reset STT hàng

    # Hiển thị dữ liệu theo năm được chọn
    selected_years = select_year_range(data)
    filtered_data = data[[data.columns[0]] + selected_years] 
    st.dataframe(filtered_data)

    # Chọn chỉ tiêu cần trực quan hóa trong selectbox
    first_column = filtered_data[filtered_data.columns[0]]
    category_row = st.selectbox("Chọn một chỉ tiêu để trực quan hóa:", first_column, key="category")
    category_data = filtered_data[first_column == category_row]
    st.dataframe(category_data)

    # Vẽ biểu đồ
    st.subheader("Trực quan hóa dữ liệu của bạn")

    chart_type = st.selectbox(
        "Choose chart type",
        ["Line Chart", "Bar Chart"], key = "chart_type"
    )

    if st.button("Vẽ biểu đồ"):
        chart_values = category_data.iloc[0, 1:].astype(float)
        chart_path = plot_chart(category_row, chart_folder_name, chart_type, selected_years, chart_values)

        bs_report = {
            "pivot_table": category_data,
            "chart_path": chart_path,
            "sheet_name": f"Sheet {st.session_state.no}",
        }

        st.session_state.balance_sheet_report.append(bs_report)
        st.session_state.no += 1

    #render_chart_from_session_state(st.session_state.balance_sheet_report, chart_folder_path)
    render_chart_from_session_state(st.session_state.balance_sheet_report)
        
    #Tính các chỉ tiêu cho phân tích cơ bản
    st.subheader("Phân tích tài chính từ dữ liệu")

 

    #chuyển cột đầu thành chữ thường
    filtered_data[filtered_data.columns[0]] = filtered_data[filtered_data.columns[0]].str.lower()

    # Các chỉ tiêu bị trùng tên trong dữ liệu sẽ được sửa tên thêm số thứ tự ở sau tên
    filtered_data_unique = filtered_data.copy()

    filtered_data_unique[filtered_data.columns[0]] = (
        filtered_data_unique
        .groupby(filtered_data.columns[0])
        .cumcount()
        .astype(str)
        .radd("_")
        .where(
            filtered_data_unique.groupby(filtered_data.columns[0]).cumcount() > 0,
            ""
        )
        + filtered_data_unique[filtered_data.columns[0]]
    )

    filtered_data_T = filtered_data_unique.set_index(filtered_data.columns[0]).T

    #tạo DataFrame chứa các chỉ tiêu tài chính
    ratio_df = calculate_financial_ratios(filtered_data_T, selected_years)
    st.dataframe(ratio_df)

    # Chọn chỉ tiêu tài chính cần trực quan hóa 
    ratios = ratio_df.columns
    ratio_visual = st.selectbox("Chọn một chỉ tiêu để trực quan hóa:", ratios, key = "ratio")
    ratio_data = ratio_df[ratio_visual]
    st.dataframe(ratio_data)

    # Vẽ biểu đồ
    st.subheader("Trực quan hóa dữ liệu của bạn")

    ratio_chart_type = st.selectbox(
        "Choose ratio chart type",
        ["Line Chart", "Bar Chart"], key="ratio_chart_type"
    )

    if st.button("Vẽ biểu đồ ratio"):
        chart_path = plot_chart(ratio_visual, chart_folder_name, ratio_chart_type, selected_years, ratio_data)

        ratio_report = {
            "pivot_table": ratio_data,
            "chart_path": chart_path,
            "sheet_name": f"Sheet {st.session_state.no}",
        }

        st.session_state.ratio_report.append(ratio_report)
        st.session_state.no += 1

    #render_chart_from_session_state(st.session_state.ratio_report, chart_folder_path)
    render_chart_from_session_state(st.session_state.ratio_report)

    #Xuất tất cả biểu đồ ra file excel
    reports = [st.session_state.balance_sheet_report, st.session_state.ratio_report]
    report_name = "report"

    # Kiểm tra nếu cả 2 reports không tồn tại trong session_state (do người dùng refresh page) thì xóa file excel report nếu đã tạo trước đó
    if st.session_state.balance_sheet_report and st.session_state.ratio_report == []:
        remove_report_file(report_name)

    if st.button("Tạo Report"):
        for report in reports:
            generate_excel_report(filtered_data, report, report_name)

        with open("report.xlsx", "rb") as file:
            excel_data = file.read()
        st.download_button(label="Tải về", data=excel_data, file_name="report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    if st.button("Reset"):
        st.session_state.clear()
        remove_report_file(report_name)
        st.rerun()






