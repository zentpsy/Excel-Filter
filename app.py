import streamlit as st
import pandas as pd
import os
from io import BytesIO
import re

st.set_page_config(page_title="Excel Filter App", layout="wide")

st.title("📊 โปรแกรมกรองข้อมูล Excel")

# โหลดข้อมูล
file_path = os.path.join("data", "all_budget.xlsx")
if not os.path.exists(file_path):
    st.error("ไม่พบไฟล์ data/all_budget.xlsx กรุณาวางไฟล์ไว้ในโฟลเดอร์ data/")
    st.stop()

df = pd.read_excel(file_path)

# ตรวจสอบคอลัมน์
required_columns = ["ลำดับ", "โครงการ", "รูปแบบงบประมาณ", "ปีงบประมาณ", "หน่วยงาน",
                    "สถานที่", "หมู่ที่", "ตำบล", "อำเภอ", "จังหวัด"]
if not all(col in df.columns for col in required_columns):
    st.error("ไฟล์ Excel ไม่มีคอลัมน์ที่ต้องการ หรือชื่อคอลัมน์ไม่ถูกต้อง")
    st.stop()

def extract_number(s):
    match = re.search(r"\d+", str(s))
    return int(match.group()) if match else float('inf')

# 1) เลือก รูปแบบงบประมาณ กับ ปีงบประมาณ ก่อน (สองตัวนี้จะส่งผลต่อกัน)
col1, col2 = st.columns(2)

# สร้าง DataFrame ชั่วคราวเพื่อใช้กรองตัวเลือกสำหรับ dropdown แรก
df_for_filter_options_1 = df.copy()

with col1:
    budget_options_base = df_for_filter_options_1["รูปแบบงบประมาณ"].dropna().unique().tolist()
    budget_options_base.sort()
    selected_budget = st.selectbox("💰 รูปแบบงบประมาณ", ["ทั้งหมด"] + budget_options_base, key="budget_select")

# กรองข้อมูลสำหรับตัวเลือกปีงบประมาณ โดยอิงจาก selected_budget
if selected_budget != "ทั้งหมด":
    df_for_filter_options_1 = df_for_filter_options_1[df_for_filter_options_1["รูปแบบงบประมาณ"] == selected_budget]

with col2:
    year_options_filtered = df_for_filter_options_1["ปีงบประมาณ"].dropna().unique().tolist()
    year_options_filtered = sorted([str(y) for y in year_options_filtered])
    selected_year = st.selectbox("📅 ปีงบประมาณ", ["ทั้งหมด"] + year_options_filtered, key="year_select")

# 2) กรองข้อมูลตาม 2 ฟิลเตอร์แรก เพื่อหา project options และ department options
# ใช้ df_temp สำหรับกรองตัวเลือกที่เหลือ
filtered_temp_for_options = df.copy()

if selected_budget != "ทั้งหมด":
    filtered_temp_for_options = filtered_temp_for_options[filtered_temp_for_options["รูปแบบงบประมาณ"] == selected_budget]

if selected_year != "ทั้งหมด":
    filtered_temp_for_options = filtered_temp_for_options[filtered_temp_for_options["ปีงบประมาณ"].astype(str) == selected_year]

col3, col4 = st.columns(2)

with col3:
    project_options = filtered_temp_for_options["โครงการ"].dropna().unique().tolist()
    project_options.sort()
    selected_project = st.selectbox("📌 โครงการ", ["ทั้งหมด"] + project_options, key="project_select")

with col4:
    department_options = filtered_temp_for_options["หน่วยงาน"].dropna().unique().tolist()
    department_options_sorted = ["ทั้งหมด"] + sorted(department_options, key=extract_number)
    selected_departments = st.multiselect("🏢 หน่วยงาน", department_options_sorted, default=["ทั้งหมด"], key="dept_select")

# 3) กรองข้อมูลทั้งหมดตาม filter
filtered_df = df.copy()

if selected_budget != "ทั้งหมด":
    filtered_df = filtered_df[filtered_df["รูปแบบงบประมาณ"] == selected_budget]

if selected_year != "ทั้งหมด":
    filtered_df = filtered_df[filtered_df["ปีงบประมาณ"].astype(str) == selected_year]

if selected_project != "ทั้งหมด":
    filtered_df = filtered_df[filtered_df["โครงการ"] == selected_project]

if "ทั้งหมด" not in selected_departments:
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
