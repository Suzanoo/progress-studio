# การแก้ปัญหา

| อาการหรือข้อความ | สาเหตุ | วิธีแก้ |
|---|---|---|
| No activities were found | XML ไม่มี Activity/Task ที่โปรแกรมอ่านได้ | Export XML ใหม่จากโปรแกรมต้นทาง หรือตรวจโครงสร้าง XML |
| Missing Activity Name | มี Activity ที่ไม่มีชื่อ | ใส่ชื่อ Activity แล้ว Export XML ใหม่ |
| Missing or invalid Plan Start | วันเริ่มว่างหรือรูปแบบอ่านไม่ได้ | แก้ Plan Start ในโปรแกรมต้นทาง |
| Missing or invalid Plan Finish | วันจบว่างหรือรูปแบบอ่านไม่ได้ | แก้ Plan Finish ในโปรแกรมต้นทาง |
| Finish is earlier than Start | วันจบก่อนวันเริ่ม | แก้ Logic ของ Schedule ก่อน Import |
| BOQ worksheet ไม่แสดง | Workbook เปิดอ่านไม่ได้หรือไม่มีชีทใช้งาน | เปิดไฟล์ด้วย Excel แล้ว Save เป็น `.xlsx` จากนั้นโหลดใหม่ |
| ข้อมูล BOQ ไม่ตรง | เลือกชีทผิด | เลือก BOQ worksheet ใหม่แล้วกด Load selected sheet |
| กด Map แล้วไม่เกิดอะไร | ยังไม่ได้เลือก Activity หรือ BOQ | ติ๊ก Activity 1 รายการ และ BOQ อย่างน้อย 1 รายการ |
| Share เกิน | Allocation รวมจะเกิน 100% | ลด Share หรือ Unmap Allocation เดิม |
| Nothing to undo | ไม่มีคำสั่งล่าสุดให้ย้อน | ใช้ Unmap กับรายการที่ต้องการแทน |
| Session ตรวจ Workbook ไม่ผ่าน | ไฟล์ถูกย้าย เปลี่ยนชื่อ หรือแก้ไข | Relink ไปยังไฟล์เดิมที่เหมือนกันทุกประการ |
| Required worksheet `main` was not found | ชีท `main` ถูกเปลี่ยนชื่อหรือไฟล์ไม่ใช่ Progress Workbook ที่ถูกต้อง | ใช้ไฟล์ที่สร้างจาก Progress Studio และห้ามเปลี่ยนชื่อ `main` |
| Percentage ไม่อัปเดต | Excel ยังไม่คำนวณสูตรใหม่ | เปิดไฟล์ด้วย Microsoft Excel รอคำนวณ แล้ว Save |
| Excel แจ้ง Repair | Workbook เก่าหรือ Table XML ถูกแก้จากภายนอก | Export ใหม่ด้วย Progress Studio 2.3.0 และหลีกเลี่ยงการแก้ Table structure |
| โปรแกรมดูเหมือนค้าง | XML หรือ BOQ มีขนาดใหญ่ | รอให้ทำงานเสร็จ และอย่ากดคำสั่งซ้ำ |

## ก่อนส่ง Issue ให้ผู้พัฒนา

เตรียมข้อมูลนี้:

```text
Progress Studio version
Windows version
XML มาจาก P6, MS Project หรือโปรแกรมใด
ข้อความ Error เต็ม ๆ
จำนวน Activity โดยประมาณ
ชื่อ BOQ worksheet
ขั้นตอนที่ทำก่อนเกิด Error
```

อย่าส่งไฟล์โครงการที่เป็นความลับ หากยังไม่ได้รับอนุญาต
