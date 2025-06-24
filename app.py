import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Excel Filter", layout="wide")

# โหลดข้อมูล
excel_path = "data/all_budget.xlsx"
try:
    df = pd.read_excel(excel_path)
except FileNotFoundError:
    st.error(f"❌ ไม่พบไฟล์ {excel_path} กรุณาวางไฟล์ในโฟลเดอร์ data/")
    st.stop()

st.title("📊 ระบบกรองข้อมูลจาก Excel")

# --- ฟังก์ชันสำหรับเรียงลำดับตัวเลขในหน่วยงาน
def extract_number(text):
    match = re.search(r'\d+', str(text))
    if match:
        return int(match.group())
    return 9999

# --- ฟิลเตอร์
col1, col2, col3, col4 = st.columns(4)

with col1:
    project_options = ["ทั้งหมด"] + sorted(df["โครงการ"].dropna().unique().tolist())
    selected_project = st.selectbox("🎯 โครงการ", project_options)

with col2:
    budget_options = ["ทั้งหมด"] + sorted(df["รูปแบบงบประมาณ"].dropna().unique().tolist())
    selected_budget = st.selectbox("💰 รูปแบบงบประมาณ", budget_options)

with col3:
    year_options = ["ทั้งหมด"] + sorted(df["ปีงบประมาณ"].dropna().unique())
    selected_year = st.selectbox("📅 ปีงบประมาณ", year_options)

with col4:
    department_options = df["หน่วยงาน"].dropna().unique().tolist()
    department_options_sorted = ["ทั้งหมด"] + sorted(department_options, key=extract_number)
    selected_department = st.selectbox("🏢 หน่วยงาน", department_options_sorted)

# --- กรองข้อมูล
filtered_df = df.copy()

if selected_project != "ทั้งหมด":
    filtered_df = filtered_df[filtered_df["โครงการ"] == selected_project]

if selected_budget != "ทั้งหมด":
    filtered_df = filtered_df[filtered_df["รูปแบบงบประมาณ"] == selected_budget]

if selected_year != "ทั้งหมด":
    filtered_df = filtered_df[filtered_df["ปีงบประมาณ"] == selected_year]

if selected_department != "ทั้งหมด":
    filtered_df = filtered_df[filtered_df["หน่วยงาน"] == selected_department]

st.markdown(f"🔎 พบข้อมูลทั้งหมด {len(filtered_df)} รายการ")
st.dataframe(filtered_df, use_container_width=True)

# --- Export Excel
def to_excel_bytes(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

if not filtered_df.empty:
    st.download_button(
        label="📥 ดาวน์โหลดข้อมูลที่กรองเป็น Excel",
        data=to_excel_bytes(filtered_df),
        file_name="filtered_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
