# ข้อกำหนด BOQ Workbook

BOQ หนึ่งไฟล์มีได้หลายชีท โปรแกรมจะให้ User เลือกชีทที่ต้องการเอง และจะไม่เดาว่าชีทใดถูกต้อง

## Column ที่ระบบออกแบบไว้รองรับ

```text
WBS-1
WBS-2
WBS-3
WBS-4
Description
Unit
Qty
Material
Labor
Amount
```

Column ที่สำคัญที่สุดต่อการ Mapping คือ **Description** และ **Amount** ส่วน WBS ใช้ช่วยจัดกลุ่มและ Filter

## เตรียม BOQ ก่อนโหลด

1. ให้ข้อมูล BOQ หลักอยู่ในตารางเดียวต่อหนึ่งชีท
2. ใช้ Header เพียงหนึ่งแถว
3. หลีกเลี่ยง Merge Cell ภายในตารางข้อมูล
4. ค่า Amount ต้องเป็นตัวเลข
5. ระวังแถว Subtotal ที่ซ้ำกับมูลค่ารายการย่อย
6. Description ควรอ่านแล้วรู้ว่าเป็นงานอะไร
7. บันทึกเป็น `.xlsx` หรือ `.xlsm`

## วิธีเลือกชีท

1. กด **Load BOQ...**
2. ดูรายชื่อใน **BOQ worksheet**
3. เลือกชีทที่มีข้อมูล BOQ จริง
4. กด **Load selected sheet**

ถ้าเลือกผิด ให้เลือกชีทใหม่แล้วโหลดอีกครั้ง

## ความหมายของ Column ในหน้าจอ Mapping

- **Amount** — มูลค่าเต็มของ BOQ Item
- **Allocated** — มูลค่าที่จัดสรรไปแล้ว
- **Remaining %** — สัดส่วนที่ยังไม่ได้จัดสรร
- **Status** — Unmapped, Partial หรือ Full
- **Mapped To** — Activity ที่รับการจัดสรร

BOQ Item หนึ่งรายการสามารถแบ่งไปหลาย Activity ได้ แต่รวมแล้วต้องไม่เกิน 100%
