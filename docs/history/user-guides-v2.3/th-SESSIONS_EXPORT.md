# การบันทึก Session และ Export

## Session คืออะไร

Session เป็นไฟล์ `.json` ที่เก็บ:

- Allocation ระหว่าง BOQ กับ Activity
- Share %
- ลายนิ้วมือของ Progress Workbook
- ลายนิ้วมือของ BOQ Workbook
- ชื่อ BOQ worksheet

Session ไม่ได้ฝัง Workbook ทั้งไฟล์ไว้ข้างใน

## บันทึก Session

1. โหลด Progress Workbook และ BOQ worksheet ให้เรียบร้อย
2. กด **Save Session...**
3. ตั้งชื่อไฟล์ เช่น `project-a.mapping.json`
4. เก็บไว้ใกล้ไฟล์โครงการ

หลัง Save ครั้งแรก เมื่อ Mapping เปลี่ยน โปรแกรมจะ Auto-save ลง Session เดิม

## กลับมาทำงานต่อ

ใช้:

- **Load Session...** เพื่อเลือกไฟล์เอง
- **Recent...** เพื่อเลือก Session ล่าสุด

โปรแกรมจะตรวจว่า Workbook ตรงกับไฟล์เดิมหรือไม่

ถ้าย้ายหรือเปลี่ยนชื่อไฟล์ โปรแกรมจะถามให้ **Relink workbook** ให้ Browse ไปหาไฟล์เดิมที่ย้ายตำแหน่งได้ แต่เนื้อหาไฟล์ต้องเหมือนเดิม หากมีการแก้ไขไฟล์ โปรแกรมจะไม่ยอมโหลด Session เพื่อป้องกัน Mapping ผิดโครงการ

## Export Workbook

1. กด **Export...**
2. ตรวจ Summary
   - Allocated amount / percentage
   - Full BOQ items
   - Partial BOQ items
   - Unmapped BOQ items
3. ถ้ายังไม่ครบ โปรแกรมจะถามว่าจะ Export Partial Mapping หรือไม่
4. เลือกชื่อไฟล์ `.xlsx`

โปรแกรมจะ:

- คัดลอก Progress Workbook อย่างปลอดภัย
- อัปเดต Amount ของ Activity ในชีท `main`
- เขียน Mapping ID, BOQ ID, Share % และ Allocated Amount
- รักษาสูตรและ Excel Table
- ตั้งค่าให้ Excel คำนวณสูตรใหม่เมื่อเปิดไฟล์

## ขั้นตอนสุดท้ายใน Microsoft Excel

1. เปิดไฟล์ที่ Export
2. รอให้ Excel คำนวณสูตร
3. ตรวจ `main`, `progress` และ `progress_table`
4. กด Save

ห้ามข้ามขั้นตอนนี้ เพราะค่า WBS และ Project Summary ใช้สูตรถ่วงน้ำหนักด้วย Amount และต้องให้ Excel คำนวณใหม่
