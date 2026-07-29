# Progress Studio 2.3.0 — คู่มือการใช้งานภาษาไทย

Progress Studio ใช้สร้าง Progress Workbook จาก Schedule XML จากนั้นนำมูลค่า BOQ มา Mapping เข้ากับ Activity บันทึกงานไว้เป็น Session และ Export เป็น Workbook ที่พร้อมใช้งานต่อ

## ลำดับการทำงานทั้งหมด

```text
Schedule XML
    ↓
สร้าง Progress Workbook
    ↓
โหลด Progress Workbook + BOQ Workbook
    ↓
Mapping มูลค่า BOQ เข้ากับ Activity
    ↓
บันทึก Mapping Session
    ↓
Export Workbook
    ↓
เปิดด้วย Microsoft Excel รอคำนวณสูตร แล้วกด Save
```

## เริ่มอ่านจากตรงนี้

- [เริ่มต้นใช้งานแบบจับมือทำ](QUICK_START.md)
- [ข้อกำหนด Schedule XML](XML_REQUIREMENTS.md)
- [ข้อกำหนด BOQ Workbook](BOQ_REQUIREMENTS.md)
- [วิธี Mapping BOQ](MAPPING_GUIDE.md)
- [การบันทึก Session และ Export](SESSIONS_EXPORT.md)
- [การแก้ปัญหา](TROUBLESHOOTING.md)

## กติกาสำคัญ

- โปรแกรมอ่านไฟล์ XML และ BOQ โดยไม่แก้ไฟล์ต้นฉบับ
- ทุก Activity ต้องมี Activity Name, Plan Start และ Plan Finish
- ถ้าขาดข้อมูลที่จำเป็นแม้แต่ Activity เดียว โปรแกรมจะหยุด Import และไม่สร้าง Workbook
- Activity ID ไม่บังคับ ถ้าไม่มีโปรแกรมจะสร้างให้
- WBS ไม่บังคับ ถ้าไม่มีโปรแกรมจะสร้างโครงสร้างแบบ Flat
- หลัง Export ต้องเปิดไฟล์ด้วย Microsoft Excel รอคำนวณสูตร และกด Save
