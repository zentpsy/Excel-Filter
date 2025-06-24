import streamlit as st
import pandas as pd
import os
from io import BytesIO

st.set_page_config(page_title="Dynamic Excel Filter", layout="wide")

# โหลดข้อมูล
file_path = os.path.join("data", "all_budget.xlsx")
if not os.path.exists(file_path):
    st.error("ไม่พบไฟล์ data/all_budget.xlsx กรุณาวางไฟล์ไว้ในโฟลเดอร์ data/")
    st.stop()

df = pd.read_excel(file_path)

# ตั้งค่าเริ่มต้น
ALL = "ทั้งหมด"

# ฟังก์ชันกรอง options สำหรับ dropdown โดยใช้เงื่อนไขที่เลือกใน dropdown อื่นๆ
def get_filtered_options(df, filter_dict, column):
    filtered_df = df.copy()
    for col, val in filter_dict.items():
        if val != ALL and col != column:
            filtered_df = filtered_df[filtered_df[col].astype(str) == val]
    options = filtered_df[column].dropna().unique().tolist()
    options.sort()
    return [ALL] + options

# เก็บสถานะเลือกจาก session_state หรือ default เป็น "ทั้งหมด"
if "selected_budget" not in st.session_state:
    st.session_state.selected_budget = ALL
if "selected_year" not in st.session_state:
    st.session_state.selected_year = ALL
if "selected_project" not in st.session_state:
    st.session_state.selected_project = ALL

# ตัวกรอง 3 ตัวนี้จะสัมพันธ์กันตลอด
filter_dict = {
    "รูปแบบงบประมาณ": st.session_state.selected_budget,
    "ปีงบประมาณ": st.session_state.selected_year,
    "โครงการ": st.session_state.selected_project,
}

col1, col2, col3 = st.columns(3)

with col1:
    budget_options = get_filtered_options(df, filter_dict, "รูปแบบงบประมาณ")
    selected_budget = st.selectbox("💰 รูปแบบงบประมาณ", budget_options, index=budget_options.index(st.session_state.selected_budget) if st.session_state.selected_budget in budget_options else 0)
    st.session_state.selected_budget = selected_budget

with col2:
    year_options = get_filtered_options(df, filter_dict, "ปีงบประมาณ")
    selected_year = st.selectbox("📅 ปีงบประมาณ", year_options, index=year_options.index(st.session_state.selected_year) if st.session_state.selected_year in year_options else 0)
    st.session_state.selected_year = selected_year

with col3:
    project_options = get_filtered_options(df, filter_dict, "โครงการ")
    selected_project = st.selectbox("📌 โครงการ", project_options, index=project_options.index(st.session_state.selected_project) if st.session_state.selected_project in project_options else 0)
    st.session_state.selected_project = selected_project

# --- หน่วยงาน (multiselect ปกติ)
import re
def extract_number(s):
    import re
    match = re.search(r"\d+", str(s))
    return int(match.group()) if match else float('inf')

department_options = df["หน่วยงาน"].dropna().unique().tolist()
department_options_sorted = [ALL] + sorted(department_options, key=extract_number)
selected_departments = st.multiselect("🏢 หน่วยงาน", department_options_sorted, default=[ALL])

# --- กรองข้อมูลตาม dropdown ทั้งหมด
filtered_df = df.copy()

if st.session_state.selected_budget != ALL:
    filtered_df = filtered_df[filtered_df["รูปแบบงบประมาณ"].astype(str) == st.session_state.selected_budget]

if st.session_state.selected_year != ALL:
    filtered_df = filtered_df[filtered_df["ปีงบประมาณ"].astype(str) == st.session_state.selected_year]

if st.session_state.selected_project != ALL:
    filtered_df = filtered_df[filtered_df["โครงการ"].astype(str) == st.session_state.selected_project]

if ALL not in selected_departments:
    filtered_df = filtered_df[filtered_df["หน่วยงาน"].isin(selected_departments)]

# แสดงจำนวนหรือแจ้งเตือน
if not filtered_df.empty:
    st.info(f"📈 พบข้อมูลทั้งหมด {len(filtered_df)} รายการ")
else:
    st.warning("⚠️ ไม่พบข้อมูลที่ตรงกับเงื่อนไขที่เลือก")

# แสดงตาราง
st.markdown("### 📄 ตารางข้อมูล")
st.dataframe(filtered_df, use_container_width=True)

# Export เป็น Excel
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
