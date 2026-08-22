# วิธี Mapping BOQ แบบละเอียด

## ทำความเข้าใจหน้าจอ

### ด้านซ้าย — Progress Activities

- ติ๊ก Activity ที่จะรับมูลค่า
- WBS แสดงไว้เพื่อให้เห็นโครงสร้าง
- Search ได้จาก Activity ID, WBS หรือ Description
- ใช้ Previous / Next เมื่อ Activity มีจำนวนมาก

### ด้านขวา — BOQ Items

- ติ๊ก BOQ หนึ่งรายการหรือหลายรายการ
- Filter ด้วย WBS-2 และ WBS-3
- Search จากคำใน Description หรือ Code
- ตรวจ Amount, Allocated, Remaining %, Status และ Mapped To

เส้นแบ่งระหว่างสองตารางสามารถลากปรับได้ โปรแกรมจะจำตำแหน่งให้

## Case 1 — Mapping เต็ม 100%

สมมุติ BOQ Item มูลค่า 100,000 ต้องลง Activity A1000 ทั้งหมด

1. ติ๊ก Activity A1000
2. ติ๊ก BOQ Item
3. ใส่ Share `100`
4. กด **Map**

ผล:

```text
Allocated = 100,000
Remaining = 0%
Status = Full
```

## Case 2 — แบ่ง BOQ Item ไปสอง Activity

สมมุติแบ่ง 60% และ 40%

Activity A1000:

1. ติ๊ก A1000
2. ติ๊ก BOQ Item
3. ใส่ Share `60`
4. กด **Map**

Activity A1010:

1. ติ๊ก A1010
2. ติ๊ก BOQ Item เดิม
3. ใส่ Share `40`
4. กด **Map**

รวม Share ต้องไม่เกิน 100%

## Case 3 — Mapping หลาย BOQ Item พร้อมกัน

เมื่อเลือก BOQ หลายรายการ ค่า Share จะใช้กับทุก BOQ Item ที่เลือก

ตัวอย่าง:

```text
เลือก BOQ 5 รายการ
Share = 25%
```

หมายความว่า BOQ ทั้ง 5 รายการจะถูกจัดสรรให้ Activity ที่เลือก รายการละ 25%

## แก้เมื่อ Mapping ผิด

- **Undo** ย้อนคำสั่ง Mapping ล่าสุด
- **Unmap** ลบ Allocation ที่เลือก
- **Clear all** ลบ Allocation ทั้งหมด โปรแกรมจะถามยืนยันก่อน

## วิธีทำงานกับ BOQ จำนวนมาก

1. Filter WBS ก่อน Search
2. ทำทีละ Work Package
3. Save Session เป็นระยะ
4. หลังโหลดไฟล์แล้ว กดลูกศรเพื่อพับ **Workbook Inputs**
5. กด **Focus Mapping** เพื่อซ่อนส่วน Generator
6. BOQ ที่ Full จะยังแสดงอยู่ เพื่อให้กลับมาแก้การตัดสินใจได้
