import streamlit as st
import pandas as pd
import os
from io import BytesIO
import re

st.set_page_config(page_title="Excel Filter App", layout="wide")

st.title("📊 Website กรองข้อมูล-งบประมาน")

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

# เก็บค่าที่เลือกจาก Dropdown แต่ละอัน
selected_budget = "ทั้งหมด"
selected_year = "ทั้งหมด"
selected_project = "ทั้งหมด"
selected_departments = ["ทั้งหมด"]

# --- ส่วนของการเลือก Filter ---
st.markdown("### 🔍 เลือกตัวกรองข้อมูล")
col1, col2 = st.columns(2)
col3, col4 = st.columns(2)

# ฟังก์ชันสำหรับดึงตัวเลือกสำหรับ Dropdown
def get_options(dataframe, column_name):
    options = dataframe[column_name].dropna().unique().tolist()
    if column_name == "ปีงบประมาณ":
        options = sorted([str(y) for y in options])
    elif column_name == "หน่วยงาน":
        options = sorted(options, key=extract_number)
    else:
        options.sort()
    return ["ทั้งหมด"] + options

# สร้าง DataFrame ชั่วคราวสำหรับการกรองตัวเลือก
filtered_df_for_options = df.copy()

# Dropdown สำหรับ 'รูปแบบงบประมาณ'
with col1:
    budget_options = get_options(filtered_df_for_options, "รูปแบบงบประมาณ")
    selected_budget = st.selectbox("💰 รูปแบบงบประมาณ", budget_options, key="budget_select")
    if selected_budget != "ทั้งหมด":
        filtered_df_for_options = filtered_df_for_options[filtered_df_for_options["รูปแบบงบประมาณ"] == selected_budget]

# Dropdown สำหรับ 'ปีงบประมาณ'
with col2:
    year_options = get_options(filtered_df_for_options, "ปีงบประมาณ")
    selected_year = st.selectbox("📅 ปีงบประมาณ", year_options, key="year_select")
    if selected_year != "ทั้งหมด":
        filtered_df_for_options = filtered_df_for_options[filtered_df_for_options["ปีงบประมาณ"].astype(str) == selected_year]

# Dropdown สำหรับ 'โครงการ'
with col3:
    project_options = get_options(filtered_df_for_options, "โครงการ")
    selected_project = st.selectbox("📌 โครงการ", project_options, key="project_select")
    if selected_project != "ทั้งหมด":
        filtered_df_for_options = filtered_df_for_options[filtered_df_for_options["โครงการ"] == selected_project]

# Dropdown สำหรับ 'หน่วยงาน' (Multiselect)
with col4:
    department_options = get_options(filtered_df_for_options, "หน่วยงาน")
    # ตรวจสอบค่า default เพื่อไม่ให้มีค่าที่ไม่ถูกต้องหลังจากการกรอง
    current_selected_departments = st.session_state.get("dept_select", ["ทั้งหมด"])
    valid_defaults = [d for d in current_selected_departments if d in department_options]
    if not valid_defaults: # ถ้าไม่มีค่าที่เคยเลือกไว้เหลืออยู่ ให้ default เป็น "ทั้งหมด"
        valid_defaults = ["ทั้งหมด"]
    
    selected_departments = st.multiselect("🏢 หน่วยงาน", department_options, default=valid_defaults, key="dept_select")
    if "ทั้งหมด" not in selected_departments:
        filtered_df_for_options = filtered_df_for_options[filtered_df_for_options["หน่วยงาน"].isin(selected_departments)]


# --- ส่วนของการกรองข้อมูลสุดท้ายเพื่อแสดงผล ---
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
def to_excel_bytes(df_to_export): # เปลี่ยนชื่อ parameter เพื่อไม่ให้ซ้ำกับ df หลัก
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_to_export.to_excel(writer, index=False)
    return output.getvalue()

if not filtered_df.empty:
    st.download_button(
        label="📥 ดาวน์โหลดข้อมูลที่กรองเป็น Excel",
        data=to_excel_bytes(filtered_df),
        file_name="filtered_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
