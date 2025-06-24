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
    st.error("ไฟล์ Excel ไม่มีคอลัมน์ที่ต้องการ")
    st.stop()

# กำหนดตัวเลือกเริ่มต้น
selected_budget = st.selectbox("💰 รูปแบบงบประมาณ", ["ทั้งหมด"])
selected_year = st.selectbox("📅 ปีงบประมาณ", ["ทั้งหมด"])
selected_project = st.selectbox("📌 โครงการ", ["ทั้งหมด"])

# อัปเดต options ตามเงื่อนไขที่เลือกไปแล้ว
filtered_df_temp = df.copy()

if selected_budget != "ทั้งหมด":
    filtered_df_temp = filtered_df_temp[filtered_df_temp["รูปแบบงบประมาณ"] == selected_budget]

if selected_year != "ทั้งหมด":
    filtered_df_temp = filtered_df_temp[filtered_df_temp["ปีงบประมาณ"].astype(str) == str(selected_year)]

if selected_project != "ทั้งหมด":
    filtered_df_temp = filtered_df_temp[filtered_df_temp["โครงการ"] == selected_project]

# อัปเดต dropdown ใหม่
col1, col2, col3 = st.columns(3)
with col1:
    budget_options = sorted(df["รูปแบบงบประมาณ"].dropna().unique().tolist())
    filtered_budget_options = [b for b in budget_options if b in filtered_df_temp["รูปแบบงบประมาณ"].unique()]
    selected_budget = st.selectbox("💰 รูปแบบงบประมาณ", ["ทั้งหมด"] + filtered_budget_options, index=(["ทั้งหมด"] + filtered_budget_options).index(selected_budget) if selected_budget in filtered_budget_options else 0)

with col2:
    year_options = sorted(df["ปีงบประมาณ"].dropna().unique().tolist())
    filtered_year_options = [y for y in year_options if y in filtered_df_temp["ปีงบประมาณ"].unique()]
    selected_year = st.selectbox("📅 ปีงบประมาณ", ["ทั้งหมด"] + [str(y) for y in filtered_year_options], index=(["ทั้งหมด"] + [str(y) for y in filtered_year_options]).index(str(selected_year)) if str(selected_year) in [str(y) for y in filtered_year_options] else 0)

with col3:
    project_options = sorted(df["โครงการ"].dropna().unique().tolist())
    filtered_project_options = [p for p in project_options if p in filtered_df_temp["โครงการ"].unique()]
    selected_project = st.selectbox("📌 โครงการ", ["ทั้งหมด"] + filtered_project_options, index=(["ทั้งหมด"] + filtered_project_options).index(selected_project) if selected_project in filtered_project_options else 0)

# กรองข้อมูลสุดท้ายอีกครั้ง
filtered_df = df.copy()
if selected_budget != "ทั้งหมด":
    filtered_df = filtered_df[filtered_df["รูปแบบงบประมาณ"] == selected_budget]
if selected_year != "ทั้งหมด":
    filtered_df = filtered_df[filtered_df["ปีงบประมาณ"].astype(str) == str(selected_year)]
if selected_project != "ทั้งหมด":
    filtered_df = filtered_df[filtered_df["โครงการ"] == selected_project]

# แสดงข้อมูล
st.markdown("### 📄 ตารางข้อมูล")
if not filtered_df.empty:
    st.info(f"📈 พบข้อมูลทั้งหมด {len(filtered_df)} รายการ")
    st.dataframe(filtered_df, use_container_width=True)
else:
    st.warning("⚠️ ไม่พบข้อมูลที่ตรงกับเงื่อนไขที่เลือก")

# Export เป็น Excel
@st.cache_data
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
