import pandas as pd
import re
import os
from tkinter import Tk, filedialog  # ใช้สำหรับเลือกไฟล์

# -------------------------------
# 🟩 ส่วนที่ 1 : เลือกไฟล์ Excel อัตโนมัติ
# -------------------------------
Tk().withdraw()  # ปิดหน้าต่างหลักของ Tkinter
input_path = filedialog.askopenfilename(
    title="เลือกไฟล์ Excel ที่ต้องการแยกข้อมูล",
    filetypes=[("Excel files", "*.xlsx *.xls")]
)

if not input_path:
    raise ValueError("❌ ไม่ได้เลือกไฟล์ Excel")

print(f"📂 เลือกไฟล์: {input_path}")

# -------------------------------
# 🟩 ส่วนที่ 2 : โหลดข้อมูล
# -------------------------------
df = pd.read_excel(input_path)
df.columns = [c.strip() for c in df.columns]

# -------------------------------
# 🟩 ส่วนที่ 3 : ตรวจหาคอลัมน์หน่วยงาน
# -------------------------------
dept_col = None
for c in df.columns:
    if "หน่วย" in c and "เกี่ยวข้อง" in c:
        dept_col = c
        break

if not dept_col:
    raise ValueError("❌ ไม่พบคอลัมน์ที่มีคำว่า 'หน่วยที่เกี่ยวข้อง' ในไฟล์ Excel")

print(f"✅ พบคอลัมน์: {dept_col}")

# -------------------------------
# 🟩 ส่วนที่ 4 : แยกข้อมูลตามหน่วยงาน
# -------------------------------
all_depts = sorted({
    d.strip()
    for v in df[dept_col].dropna().astype(str)
    for d in re.split(r"[,/ ]+", v)
    if d.strip()
})

print("📋 หน่วยงานที่พบ:", all_depts)

# -------------------------------
# 🟩 ส่วนที่ 5 : สร้างไฟล์ Excel หลายชีต
# -------------------------------
folder = os.path.dirname(input_path)
output_path = os.path.join(folder, "RM_by_department.xlsx")

with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    for dept in all_depts:
        dept_df = df[df[dept_col].astype(str).str.contains(rf"\b{dept}\b", na=False)]
        if not dept_df.empty:
            dept_df.to_excel(writer, index=False, sheet_name=dept[:31])

print(f"\n✅ สร้างไฟล์สำเร็จ: {output_path}")
