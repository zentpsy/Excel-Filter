# 📊 Streamlit Excel Filter App

แอปพลิเคชันที่แสดงข้อมูลจากไฟล์ Excel พร้อมระบบกรอง และสามารถดาวน์โหลดข้อมูลที่กรองออกมาเป็นไฟล์ใหม่

## 🗂️ โครงสร้าง

```
.
├── app.py
├── data/
│   └── all_budget.xlsx   ← วางไฟล์ Excel ของคุณที่นี่
├── requirements.txt
└── README.md
```

## 🚀 วิธีใช้งาน

### รันบนเครื่อง

1. ติดตั้งไลบรารี:
   ```bash
   pip install -r requirements.txt
   ```
2. รันแอป:
   ```bash
   streamlit run app.py
   ```

### Deploy ออนไลน์ (Streamlit Cloud)

1. Push โครงสร้างนี้ขึ้น GitHub
2. เข้า [Streamlit Cloud](https://streamlit.io/cloud) แล้วเชื่อม GitHub
3. เลือก repo นี้ ระบบจะ build และแสดงแอปให้โดยอัตโนมัติ

## ✨ ฟีเจอร์

- กรองตาม โครงการ / ปีงบฯ / หน่วยงาน / รูปแบบงบ
- แสดงข้อมูลแบบตารางกว้าง
- ดาวน์โหลดข้อมูลที่กรองออกมาเป็น Excel
