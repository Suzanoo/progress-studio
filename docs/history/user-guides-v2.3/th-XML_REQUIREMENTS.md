# ข้อกำหนด Schedule XML

Progress Studio ไม่บังคับว่า XML ต้องมาจาก Primavera P6 เท่านั้น สามารถมาจาก Microsoft Project หรือโปรแกรมอื่นได้ ถ้าโปรแกรมสามารถอ่านและแปลง Activity ให้ตรงตาม Contract

## ข้อมูลที่ต้องมีจริง ๆ ในทุก Activity

| ข้อมูล | กติกา |
|---|---|
| Activity Name | ต้องมีและห้ามเป็นค่าว่าง |
| Plan Start | ต้องเป็นวันที่ที่อ่านได้ |
| Plan Finish | ต้องเป็นวันที่ที่อ่านได้ และห้ามก่อน Plan Start |

ถ้า Activity ใดขาดข้อมูลเหล่านี้ โปรแกรมจะ:

```text
แจ้งรายการข้อผิดพลาด
↓
หยุด Import
↓
ไม่สร้าง Workbook บางส่วน
```

## ข้อมูลที่ไม่บังคับ

| ข้อมูล | ถ้าไม่มี โปรแกรมทำอย่างไร |
|---|---|
| Activity ID | สร้างรหัส เช่น `ACT-000001` ให้โดยอัตโนมัติ |
| WBS / hierarchy | สร้างโครงสร้าง Project แบบ Flat |
| Calendar | ไม่จำเป็นต่อการสร้าง Workbook |
| Relationships | ไม่จำเป็น |
| Duration | คำนวณจาก Start และ Finish ได้ |
| Actual Start / Finish | ไม่บังคับ |
| Percent Complete | ไม่บังคับ |
| Resource / Activity Code | ไม่บังคับ |

## ตัวอย่าง XML ที่ผ่าน

```xml
<Task>
  <Name>งานขุดดิน</Name>
  <Start>2026-03-01T08:00:00</Start>
  <Finish>2026-03-10T17:00:00</Finish>
</Task>
```

## ตัวอย่าง XML ที่ไม่ผ่าน

ไม่มี Activity Name:

```xml
<Task>
  <Start>2026-03-01</Start>
  <Finish>2026-03-10</Finish>
</Task>
```

Finish ก่อน Start:

```xml
<Task>
  <Name>งานขุดดิน</Name>
  <Start>2026-03-10</Start>
  <Finish>2026-03-01</Finish>
</Task>
```

## รูปแบบวันที่ที่แนะนำ

ใช้ ISO Date จะปลอดภัยที่สุด:

```text
2026-03-01
2026-03-01T08:00:00
2026-03-01T08:00:00+07:00
```

หลีกเลี่ยงวันที่กำกวม เช่น `03/04/2026` เพราะอาจหมายถึง 3 เมษายน หรือ 4 มีนาคม
