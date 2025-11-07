import pandas as pd
from datetime import datetime
import re
import os
from openpyxl.styles import Alignment
from openpyxl.worksheet.page import PageMargins
from tkinter import Tk, filedialog  # ใช้สำหรับเลือกไฟล์

print(f"📂 กรุณารอสักครู่....")
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
dnow = datetime.now()
d_formatted = dnow.strftime("%Y%m%d%H%M%S")
print_date = dnow.now().strftime("%d/%m/%Y %H:%M")

folder = os.path.dirname(".")
output_path = os.path.join(folder, f"cnh_rm_dep_{d_formatted}.xlsx")

with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    for dept in all_depts:
        dept_df = df[df[dept_col].astype(str).str.contains(rf"\b{dept}\b", na=False)]
        if not dept_df.empty:
            dept_df.to_excel(writer, index=False, sheet_name=dept[:31])

    workbook = writer.book
    for sheet_name in workbook.sheetnames:
        ws = workbook[sheet_name]
        # warp text
        for row in ws.iter_rows():
            for cell in row:
                col_idx = cell.column  # เริ่มที่ 1
                if col_idx in [1, 2]: # 2 คอลัมแรก ไม่ wraptext
                    cell.alignment = Alignment(wrapText=False, vertical="top")
                else:
                    cell.alignment = Alignment(wrapText=True, vertical="top")

        for column in ws.columns:
            max_length = 0
            col_letter = column[0].column_letter
            col_idx = column[0].column
            for cell in column:
                value = str(cell.value) if cell.value else ""
                max_length = max(max_length, len(value))

            base_width = min(max_length + 2, 18)
            if 3 <= col_idx <= 5: # คอลัม 3,4,5
                base_width *= 1.9  # กว้างขึ้น %
            
            if col_idx > 5 :
                base_width = min(max_length + 2, 16)

            ws.column_dimensions[col_letter].width = base_width

        # ตั้งค่าหน้ากระดาษแนวนอน / ขนาด A4
        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
        ws.page_setup.paperSize = ws.PAPERSIZE_A4

        # ตั้งค่าขอบกระดาษ
        ws.page_margins = PageMargins(left=0.1, right=0.1, top=0.5, bottom=0.5)

        # พิมพ์พอดีกับความกว้าง 1 หน้า
        ws.page_setup.fitToPage = True      # เปิดโหมด Fit to Page
        ws.page_setup.fitToWidth = 1        # พอดีกับความกว้างหน้าเดียว
        ws.page_setup.fitToHeight = 0       # ไม่จำกัดความสูง
        ws.page_setup.scale = None          # ปิด scale manual เพื่อใช้ fitToPage แทน

        ws.HeaderFooter.leftHeader = ""
        ws.HeaderFooter.centerHeader = f"&Bแผนก: {sheet_name}&B"
        ws.HeaderFooter.rightHeader = f"วันที่พิมพ์: {print_date}"

        ws.HeaderFooter.leftFooter = "&F"
        ws.HeaderFooter.centerFooter = ""
        ws.HeaderFooter.rightFooter = "หน้า &P จาก &N"

        # ✅ ตั้งให้พิมพ์หัวตาราง (row 1) ซ้ำทุกหน้า
        ws.print_title_rows = "1:1"

print(f"\n✅ สร้างไฟล์สำเร็จ: {output_path}")
