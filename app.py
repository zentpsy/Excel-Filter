import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="ข้อมูลโครงการ", layout="wide")
st.title("📊 แสดงข้อมูลโครงการจาก Excel พร้อมฟิลเตอร์")

excel_file_path = "data/all_budget.xlsx"

# โหลดข้อมูล
if os.path.exists(excel_file_path):
    try:
        df = pd.read_excel(excel_file_path)
        expected_columns = [
            "ลำดับ", "โครงการ", "รูปแบบงบประมาณ", "ปีงบประมาณ",
            "หน่วยงาน", "สถานที่", "หมู่ที่", "ตำบล", "อำเภอ", "จังหวัด"
        ]
        if all(col in df.columns for col in expected_columns):
            for col in ["หมู่ที่", "ลำดับ"]:
                df[col] = df[col].fillna('').astype(str)

            col1, col2 = st.columns(2)
            with col1:
                selected_project = st.selectbox("🧾 เลือกโครงการ", ["ทั้งหมด"] + sorted(df["โครงการ"].dropna().unique()))
                selected_budget = st.selectbox("💰 เลือกรูปแบบงบประมาณ", ["ทั้งหมด"] + sorted(df["รูปแบบงบประมาณ"].dropna().unique()))
            with col2:
                selected_year = st.selectbox("📅 เลือกปีงบประมาณ", ["ทั้งหมด"] + sorted(df["ปีงบประมาณ"].dropna().unique()))
                depts = df["หน่วยงาน"].dropna().unique().tolist()
                numeric_depts = []
                text_depts = []
                for d in depts:
                    try:
                        numeric_depts.append((float(str(d)), d))
                    except:
                        text_depts.append(d)
                numeric_depts.sort()
                text_depts.sort()
                sorted_depts = [d[1] for d in numeric_depts] + text_depts

                selected_departments = st.multiselect(
                    "📍 เลือกหน่วยงาน (เลือกได้หลายค่า)",
                    ["ทั้งหมด"] + sorted_depts,
                    default=["ทั้งหมด"]
                )

            filtered_df = df.copy()
            if selected_project != "ทั้งหมด":
                filtered_df = filtered_df[filtered_df["โครงการ"] == selected_project]
            if selected_budget != "ทั้งหมด":
                filtered_df = filtered_df[filtered_df["รูปแบบงบประมาณ"] == selected_budget]
            if selected_year != "ทั้งหมด":
                filtered_df = filtered_df[filtered_df["ปีงบประมาณ"] == selected_year]
            if selected_departments and "ทั้งหมด" not in selected_departments:
                filtered_df = filtered_df[filtered_df["หน่วยงาน"].isin(selected_departments)]

            st.info(f"📈 พบข้อมูลทั้งหมด {len(filtered_df)} รายการ")
            st.dataframe(filtered_df, use_container_width=True, height=600)

            @st.cache_data
            def to_excel(df):
                return df.to_excel(index=False, engine="openpyxl")

            if not filtered_df.empty:
                st.download_button(
                    label="📥 ดาวน์โหลดข้อมูลที่กรองเป็น Excel",
                    data=to_excel(filtered_df),
                    file_name="filtered_data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        else:
            st.error("❌ คอลัมน์ไม่ครบ โปรดตรวจสอบ: " + ", ".join(expected_columns))
    except Exception as e:
        st.error(f"เกิดข้อผิดพลาด: {e}")
else:
    st.error(f"ไม่พบไฟล์ `{excel_file_path}` กรุณาวางไฟล์ในโฟลเดอร์ `data/`")
