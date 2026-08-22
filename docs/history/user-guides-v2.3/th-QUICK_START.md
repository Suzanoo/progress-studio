# เริ่มต้นใช้งานแบบจับมือทำ

## ขั้นตอนที่ 1 — ติดตั้งและเปิดโปรแกรม

เปิด PowerShell ที่โฟลเดอร์โปรเจกต์ แล้วรัน:

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python desktop.py
```

หน้าต่าง Progress Studio จะเปิดแบบเต็มหน้าจอ

## ขั้นตอนที่ 2 — สร้าง Progress Workbook จาก XML

1. ดูส่วน **Project input**
2. ที่ช่อง **Schedule XML** กด **Browse...**
3. เลือกไฟล์ XML ของโครงการ
4. เลือก **Weekly cutoff day** หรือวันตัดรอบประจำสัปดาห์
5. ช่อง **Fallback amount / activity** ใช้เฉพาะกรณี XML ไม่มีข้อมูล Amount และต้องการใส่มูลค่าชั่วคราวให้ทุก Activity
6. เลือก **Plan distribution**
   - `Auto` ให้โปรแกรมเลือกการกระจาย
   - `Flat` กระจายเท่ากัน
   - `Front` น้ำหนักมากช่วงต้น
   - `Back` น้ำหนักมากช่วงท้าย
   - `Bell` น้ำหนักมากช่วงกลาง
7. กด **Create Progress Workbook**
8. รอจนโปรแกรมแจ้งว่าเสร็จ
9. กด **Open output workbook** หรือ **Open output folder**

ไฟล์ XML ต้นฉบับจะไม่ถูกแก้ไข

## ขั้นตอนที่ 3 — โหลดไฟล์สำหรับ Mapping

เปิดแท็บ **Amount Mapping**

1. กด **Load Progress...** แล้วเลือก Progress Workbook ที่เพิ่งสร้าง
2. กด **Load BOQ...** แล้วเลือก BOQ Workbook
3. เลือกชีทจากช่อง **BOQ worksheet**
4. กด **Load selected sheet**

เมื่อโหลดสำเร็จ ด้านซ้ายจะแสดง Activity และด้านขวาจะแสดง BOQ Items

## ขั้นตอนที่ 4 — Mapping BOQ เข้ากับ Activity

1. ติ๊ก Activity ที่ต้องการทางด้านซ้าย 1 รายการ
2. ติ๊ก BOQ Item ทางด้านขวา 1 รายการหรือหลายรายการ
3. ใส่ **Share = 100** เมื่อต้องการจัดสรรเต็มจำนวน
4. ถ้าต้องการแบ่ง BOQ ไปหลาย Activity ให้ใส่ Share ต่ำกว่า 100
5. กด **Map**
6. ตรวจสอบ Summary ด้านบน

```text
Mapped / Total | Remaining | จำนวน BOQ ที่ Mapping แล้ว / ทั้งหมด
```

กด **Undo** เมื่อต้องการย้อนคำสั่งล่าสุด หรือกด **Unmap** เพื่อลบ Mapping ที่เลือก

## ขั้นตอนที่ 5 — บันทึก Session

1. กด **Save Session...**
2. ตั้งชื่อไฟล์ `.json`
3. เก็บไฟล์ Session ไว้ใกล้กับ Progress Workbook และ BOQ Workbook

หลังจาก Save ครั้งแรก การแก้ Mapping ครั้งต่อไปจะ Auto-save ลงไฟล์เดิม

## ขั้นตอนที่ 6 — Export

1. กด **Export...**
2. ตรวจสอบ Allocated, Full, Partial และ Unmapped
3. ถ้ายัง Mapping ไม่ครบ โปรแกรมจะถามว่าจะ Export แบบ Partial หรือไม่
4. เลือกชื่อไฟล์ `.xlsx`
5. เปิดไฟล์ที่ Export ด้วย Microsoft Excel
6. รอให้ Excel คำนวณสูตรจนเสร็จ
7. ตรวจสอบข้อมูลแล้วกด Save

ขั้นตอนเปิดและ Save ด้วย Excel สำคัญ เพราะชีท `progress` และ `progress_table` ใช้สูตรคำนวณจากข้อมูลใน `main`
