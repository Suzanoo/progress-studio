"""Presentation-only resources. English IDs are stable; values never enter engines."""

GUIDE = {
    "preview.summary": ("Ordinary: {ordinary} | Positive: {positive} | Zero: {zero} | Milestones: {milestones} (zero weight) | Valid-value total: {total:,.2f}", "กิจกรรมทั่วไป: {ordinary} | บวก: {positive} | ศูนย์: {zero} | Milestone: {milestones} (น้ำหนักศูนย์) | รวมค่าที่ถูกต้อง: {total:,.2f}"),
    "preview.blocked": ("Create blocked: {count} issue(s). {error}", "ยังสร้างไม่ได้: พบ {count} ปัญหา รายละเอียด: {error}"),
    "preview.precision": ("Values retain workbook numeric precision; two decimals are display only.", "ค่าคงความละเอียดตามตัวเลขของ Workbook ทศนิยมสองตำแหน่งใช้แสดงผลเท่านั้น"),
    "home.title": ("Your progress workbook", "เริ่มต้นใช้งาน Progress Studio"),
    "home.flow": ("Select XML  →  Create Excel  →  Update Progress", "เลือก XML  →  สร้าง Excel  →  บันทึกความก้าวหน้า"),
    "home.create": ("Create workbook", "สร้าง Workbook"),
    "home.create_help": ("Create an Excel progress workbook from P6 / MSP XML.", "สร้างไฟล์ Excel จาก P6 / MSP XML"),
    "home.rebuild": ("Rebuild workbook", "อัปเดต Workbook"),
    "home.rebuild_help": ("Refresh views from your saved Excel workbook.", "อัปเดตมุมมองจากไฟล์ Excel ที่บันทึกแล้ว"),
    "home.mapping": ("BOQ Mapping", "จับคู่ BOQ"),
    "home.mapping_help": ("Allocate BOQ amounts to activities.", "เชื่อมมูลค่า BOQ กับกิจกรรม"),
    "home.payment": ("Payment", "งวดงาน"),
    "home.payment_help": ("Prepare requirements and Payment Breakdown.", "จัดเตรียมข้อมูลและสรุปผลงานงวดงาน"),
    "home.optional": ("Optional", "ทางเลือก"),
    "guide.title": ("Help / Quick Guide", "วิธีใช้"),
    "guide.view": ("View guide", "ดูวิธีใช้"),
    "guide.back": ("Back to Home", "กลับหน้าหลัก"),
    "guide.first": ("First time — Create workbook", "ครั้งแรก — สร้าง Workbook"),
    "guide.first.1": ("Select XML", "เลือก XML"),
    "guide.first.1.help": ("Export from P6 or MSP.", "ส่งออกจาก P6 หรือ MSP"),
    "guide.first.2": ("Configure and Create", "ตั้งค่าและสร้างไฟล์"),
    "guide.first.2.help": ("Weight / Cutoff / Distribution", "น้ำหนัก / วันตัดรอบ / การกระจายแผน"),
    "guide.first.3": ("Open Excel", "เปิด Excel"),
    "guide.first.3.help": ("Review plan and S-curve.", "ตรวจแผนงานและ S-curve"),
    "guide.cycle": ("Each reporting cycle — Update progress", "แต่ละรอบ — อัปเดตความก้าวหน้า"),
    "guide.cycle.1": ("Enter progress in Excel", "กรอกผลงานใน Excel"),
    "guide.cycle.1.help": ("Use editable cells in main.", "ใช้ช่องที่เปิดให้แก้ไขใน main"),
    "guide.cycle.2": ("F9 and Save", "F9 แล้ว Save"),
    "guide.cycle.2.help": ("Recalculate and save the workbook.", "คำนวณและบันทึกไฟล์"),
    "guide.cycle.3": ("Rebuild when required", "Rebuild เมื่อจำเป็น"),
    "guide.cycle.3.help": ("Select the saved .xlsx and Target.", "เลือกไฟล์ .xlsx ที่บันทึกแล้วและงานที่จะอัปเดต"),
    "guide.rebuild_note": ("Rebuild uses the saved workbook, not a new XML import. Use it to refresh generated snapshots or affected structural data; it is not mandatory after every progress edit. Progress / Payment retain Snapshot and Live options. F9 does not regenerate snapshots or structural BAC.", "Rebuild ใช้ Workbook ที่บันทึกแล้ว ไม่ใช่นำเข้า XML ใหม่ ใช้เมื่ออัปเดตมุมมองที่สร้างไว้หรือข้อมูลโครงสร้างที่เกี่ยวข้อง ไม่ต้องทำหลังกรอกผลงานทุกครั้ง Progress / Payment เลือก Snapshot หรือ Live ได้ F9 ไม่สร้าง Snapshot หรือปรับ BAC เชิงโครงสร้างใหม่"),
    "guide.optional": ("Optional tools: BOQ Mapping • Payment • Earned Value", "เครื่องมือเพิ่มเติม: BOQ Mapping • Payment • Earned Value"),
    "guide.ev_note": ("Earned Value is in Rebuild and uses Live formulas. Ordinary progress edits recalculate with F9 / Save; BAC or monetary source changes require EV refresh.", "Earned Value อยู่ใน Rebuild และใช้สูตร Live การแก้ผลงานปกติคำนวณด้วย F9 / Save หากเปลี่ยน BAC หรือแหล่งมูลค่า ต้องอัปเดต EV"),
    "language.help": ("Choose ENG or THA at the top right. The preference applies on the next application launch. Workbook names, formulas and source fields are unchanged.", "เลือก ENG หรือ THA ที่มุมบนขวา ภาษาที่เลือกมีผลเมื่อเปิดแอปครั้งถัดไป ชื่อชีต สูตร และฟิลด์ต้นทางไม่เปลี่ยน"),
    "language.restart": ("Language saved. It will apply the next time Progress Studio is opened. Your current work remains open.", "บันทึกภาษาแล้ว จะใช้เมื่อเปิด Progress Studio ครั้งถัดไป งานปัจจุบันยังเปิดอยู่"),
    "language.save_failed": ("Could not save language preference: {error}", "บันทึกภาษาที่เลือกไม่ได้: {error}"),
    "error.details": ("The operation could not be completed. Details: {error}", "ดำเนินการไม่สำเร็จ รายละเอียด: {error}"),
    "status.not_ready": ("Not ready — {error}", "ยังไม่พร้อม — {error}"),
    "status.ev_ready": ("Ready • {count:,} activities • {source}", "พร้อม • {count:,} กิจกรรม • {source}"),
    "status.created": ("Created {name}", "สร้างไฟล์แล้ว: {name}"),
}

PAIRS = dict(line.split(" => ", 1) for line in '''Home => หน้าหลัก
Create => สร้างไฟล์
Mapping => จับคู่ BOQ
Payment => งวดงาน
Rebuild => อัปเดต Workbook
Settings => ตั้งค่า
Help / Quick Guide => วิธีใช้
No Amount field selected => ยังไม่ได้เลือกฟิลด์ Amount
Ready => พร้อมใช้งาน
Select an XML schedule file to begin. => เลือกไฟล์แผนงาน XML เพื่อเริ่มต้น
Mapping Workspace => จับคู่ BOQ กับกิจกรรม
Open Project... => เปิดโครงการ...
Recent Projects => โครงการล่าสุด
Save => บันทึก
Save As... => บันทึกเป็น...
Save As => บันทึกเป็น
Exit => ออกจากโปรแกรม
File => ไฟล์
Undo => ย้อนกลับ
Map Selection => จับคู่รายการที่เลือก
Unmap Selection => ยกเลิกการจับคู่ที่เลือก
Edit => แก้ไข
Toggle Sidebar => แสดง/ซ่อนเมนูด้านข้าง
Focus Mapping => ขยายพื้นที่จับคู่ BOQ
View => มุมมอง
Import Workspace => สร้าง Workbook
Payment Workspace => งวดงาน
Rebuild Workspace => อัปเดต Workbook
Tools => เครื่องมือ
About Progress Studio => เกี่ยวกับ Progress Studio
Help => วิธีใช้
Create Progress Workbook => สร้าง Progress Workbook
Open output workbook => เปิด Workbook ผลลัพธ์
Open output folder => เปิดโฟลเดอร์ผลลัพธ์
Select Schedule XML => เลือกแผนงาน XML
Running => กำลังทำงาน
Starting pipeline... => กำลังเริ่มสร้างไฟล์...
Completed => เสร็จสมบูรณ์
Progress workbook is ready. => Progress Workbook พร้อมใช้งาน
Failed => ไม่สำเร็จ
Pipeline stopped. Review the activity log. => หยุดการทำงาน กรุณาดูรายละเอียดในบันทึกการทำงาน
Progress Studio => Progress Studio
Financial Forecast => Financial Forecast
Wait for the finance operation to finish before closing. => กรุณารองาน Finance เสร็จก่อนปิดโปรแกรม
Input selected. Review options and create the workbook. => เลือกไฟล์แล้ว ตรวจตัวเลือกและสร้าง Workbook
Working... => กำลังทำงาน...
Export Mapped Workbook => ส่งออก Workbook ที่จับคู่แล้ว
Go to Mapping => ไปหน้าจับคู่ BOQ
Activity Log => บันทึกการทำงาน
Schedule XML => แผนงาน XML
Browse... => เลือกไฟล์...
Weekly cutoff day => วันตัดรอบรายสัปดาห์
Weight basis => ฐานน้ำหนัก
Select Amount field / Preview... => เลือกฟิลด์ Amount / ดูตัวอย่าง...
Plan distribution => การกระจายแผน
XML Amount field => ฟิลด์ Amount ใน XML
Invalid input => ข้อมูลไม่ถูกต้อง
Open failed => เปิดไม่สำเร็จ
No Progress workbook loaded => ยังไม่ได้โหลด Progress Workbook
No BOQ workbook loaded => ยังไม่ได้โหลด BOQ Workbook
Load both workbooks to begin mapping. => โหลด Workbook ทั้งสองไฟล์เพื่อเริ่มจับคู่
Mapped 0.00 / 0.00 | Remaining 0.00 | Items 0/0 => จับคู่ 0.00 / 0.00 | คงเหลือ 0.00 | รายการ 0/0
Rows 0-0 of 0 => แถว 0-0 จาก 0
Selected 0 items | 0.00 => เลือก 0 รายการ | 0.00
Select a BOQ item to view all mapped activities. => เลือกรายการ BOQ เพื่อดูกิจกรรมที่จับคู่
Project: unsaved => โครงการ: ยังไม่บันทึก
No BOQ worksheet loaded => ยังไม่ได้โหลดชีต BOQ
Mapping Inputs => ข้อมูลสำหรับจับคู่
Load selected sheet => โหลดชีตที่เลือก
Arrange => จัดลำดับ
Move Up => เลื่อนขึ้น
Move Down => เลื่อนลง
Indent => ลดระดับ
Outdent => เพิ่มระดับ
Move to WBS... => ย้ายไป WBS...
Undo tree edit => ย้อนการแก้โครงสร้าง
Redo tree edit => ทำซ้ำการแก้โครงสร้าง
Tree View => มุมมองโครงสร้าง
Expand all => ขยายทั้งหมด
Collapse all => ยุบทั้งหมด
Select edited Progress workbook => เลือก Progress Workbook ที่แก้ไขแล้ว
Session: not saved => งานปัจจุบัน: ยังไม่บันทึก
Select Progress workbook => เลือก Progress Workbook
Select BOQ workbook => เลือก BOQ Workbook
Add WBS => เพิ่ม WBS
WBS code: => รหัส WBS:
WBS name: => ชื่อ WBS:
Add Activity => เพิ่มกิจกรรม
Activity ID (must be unique): => Activity ID (ต้องไม่ซ้ำ):
Activity name: => ชื่อกิจกรรม:
Name: => ชื่อ:
Move progress node => ย้ายรายการแผนงาน
Amount Mapping => จับคู่มูลค่า
Save Progress Studio project => บันทึกโครงการ Progress Studio
Open Progress Studio project => เปิดโครงการ Progress Studio
Export mapped Progress workbook => ส่งออก Progress Workbook ที่จับคู่แล้ว
Load a BOQ workbook first. => กรุณาโหลด BOQ Workbook ก่อน
Select a BOQ worksheet. => กรุณาเลือกชีต BOQ
Progress tree => โครงสร้างแผนงาน
Load a Progress workbook first. => กรุณาโหลด Progress Workbook ก่อน
Select a WBS first. => กรุณาเลือก WBS ก่อน
Select a WBS or Activity to edit. => เลือก WBS หรือกิจกรรมที่จะแก้ไข
Select a WBS or Activity to delete. => เลือก WBS หรือกิจกรรมที่จะลบ
Delete progress node => ลบรายการแผนงาน
Select a WBS or Activity to move. => เลือก WBS หรือกิจกรรมที่จะย้าย
No destination WBS is available. => ไม่มี WBS ปลายทาง
Nothing to undo. => ไม่มีรายการให้ย้อนกลับ
Clear every mapping? You can use Undo immediately after this command. => ล้างการจับคู่ทั้งหมดหรือไม่? สามารถกดย้อนกลับได้ทันทีหลังคำสั่งนี้
Project: waiting for both workbooks => โครงการ: รอ Workbook ทั้งสองไฟล์
Auto-saved => บันทึกอัตโนมัติแล้ว
No recent projects were found. => ไม่พบโครงการล่าสุด
Export partial mapping? => ส่งออกการจับคู่บางส่วนหรือไม่?
Workbook Inputs => ไฟล์ Workbook ต้นทาง
Progress workbook => Progress Workbook
Load Progress... => โหลด Progress...
BOQ workbook => BOQ Workbook
Load BOQ... => โหลด BOQ...
BOQ worksheet => ชีต BOQ
Progress Activities => กิจกรรมแผนงาน
Delete => ลบ
BOQ Items => รายการ BOQ
WBS-2 => WBS-2
WBS-3 => WBS-3
Select page => เลือกทั้งหน้า
Select all filtered => เลือกทั้งหมดที่กรอง
Clear selection => ยกเลิกการเลือก
Mapped activities => กิจกรรมที่จับคู่
Share => สัดส่วน
Map => จับคู่
Unmap => ยกเลิกจับคู่
Clear all => ล้างทั้งหมด
Share applies to every selected BOQ item. => สัดส่วนนี้ใช้กับทุกรายการ BOQ ที่เลือก
Search => ค้นหา
Clear => ล้าง
Previous => ก่อนหน้า
Next => ถัดไป
Move => ย้าย
Cancel => ยกเลิก
Project: auto-save failed => โครงการ: บันทึกอัตโนมัติไม่สำเร็จ
Select a project to continue: => เลือกโครงการเพื่อทำงานต่อ:
Open => เปิด
Relink workbook => เชื่อมไฟล์ Workbook ใหม่
Export mapped workbook => ส่งออก Workbook ที่จับคู่แล้ว
Export complete => ส่งออกสำเร็จ
Select XML Amount field => เลือกฟิลด์ Amount ใน XML
Use selected field => ใช้ฟิลด์ที่เลือก
Select a numeric custom field. Values become monetary weights; milestones remain zero. => เลือกฟิลด์ตัวเลขเพื่อใช้เป็นน้ำหนักตามมูลค่า Milestone ยังคงมีน้ำหนักเป็นศูนย์
Select an exported Progress Studio workbook. => เลือก Workbook ที่ส่งออกจาก Progress Studio
Default is calculated from Project Start / Finish. => ค่าเริ่มต้นคำนวณจากวันเริ่มและสิ้นสุดโครงการ
Prepare or reconcile the persistent Payment Input sheet. => เตรียมหรือปรับข้อมูลในชีต Payment Input
Build Payment-Breakdown from repeated exact Activity Names in main. => สร้าง Payment-Breakdown จากกิจกรรมชื่อเหมือนกันใน main
Prepare Payment Input => เตรียม Payment Input
Open Result => เปิดผลลัพธ์
Build Payment Breakdown => สร้าง Payment Breakdown
Select Progress Studio workbook => เลือก Progress Studio Workbook
Save workbook with Payment Input => บันทึก Workbook พร้อม Payment Input
Preparing Payment Input... => กำลังเตรียม Payment Input...
Payment Input preparation failed. => เตรียม Payment Input ไม่สำเร็จ
Save workbook with Payment Breakdown => บันทึก Workbook พร้อม Payment Breakdown
Building Payment Breakdown... => กำลังสร้าง Payment Breakdown...
Payment Breakdown build failed. => สร้าง Payment Breakdown ไม่สำเร็จ
Select a valid workbook first. => กรุณาเลือก Workbook ที่ถูกต้องก่อน
Payment periods must be between 1 and 120. => จำนวนงวดต้องอยู่ระหว่าง 1 ถึง 120
Prepare Payment Input or derive Payment-Breakdown here. Use Rebuild when you want to regenerate Payment lines. => เตรียม Payment Input หรือสร้าง Payment-Breakdown ที่นี่ ใช้ Rebuild เมื่อต้องการสร้างเส้น Payment ใหม่
Workbook => ไฟล์ Workbook
Select the workbook that will receive Payment Input or a derived Payment-Breakdown snapshot. => เลือก Workbook ที่จะเพิ่ม Payment Input หรือผลสรุป Payment-Breakdown
Existing user percentages are preserved by Activity ID. New Activities receive suggested fake requirements. Payment Date is not an input. => รักษาเปอร์เซ็นต์เดิมตาม Activity ID กิจกรรมใหม่ได้ค่าแนะนำเบื้องต้น โดยไม่ต้องกรอก Payment Date
Payment periods => จำนวนงวดงาน
Payment Breakdown => สรุปผลงานงวดงาน
Derive repeated exact Activity Names from current main. Each source Activity keeps its own progress first; the combined row is Amount-weighted. => รวมกิจกรรมชื่อเหมือนกันจาก main โดยแต่ละกิจกรรมใช้ผลงานของตนเอง แล้วถ่วงน้ำหนักแถวรวมด้วย Amount
Review the period count before preparing. => ตรวจจำนวนงวดก่อนเตรียมข้อมูล
Payment periods must be a whole number. => จำนวนงวดต้องเป็นจำนวนเต็ม
Select a workbook to analyze. => เลือก Workbook เพื่อตรวจสอบ
main is preserved and used as the source of truth. => รักษาข้อมูลใน main และใช้เป็นข้อมูลต้นทาง
Select a workbook to check Earned Value readiness. => เลือก Workbook เพื่อตรวจความพร้อมของ Earned Value
Select EV monetary source. Activity Amount uses Allocated Contract Value. => เลือกแหล่งมูลค่า EV โดย Activity Amount ใช้มูลค่าสัญญาที่จัดสรรให้กิจกรรม
Progress => ความก้าวหน้า
Prepare monetary inputs => เตรียมข้อมูลมูลค่า
Select workbook to rebuild => เลือก Workbook ที่จะอัปเดต
Save rebuilt workbook => บันทึก Workbook ที่อัปเดต
Prepare EV monetary inputs => เตรียมข้อมูลมูลค่า EV
Earned Value => Earned Value
Open the new workbook and fill EV Monetary Inputs with allocated Contract Values. Mark milestones and enter their dates. Save, select this workbook, and refresh EV. => เปิด Workbook ใหม่ กรอกมูลค่าสัญญาที่จัดสรรใน EV Monetary Inputs ระบุ Milestone และวันที่ แล้วบันทึก เลือกไฟล์นี้และอัปเดต EV
Save Earned Value workbook => บันทึก Workbook Earned Value
Generating Earned Value view... => กำลังสร้างมุมมอง Earned Value...
Rebuild failed. => อัปเดต Workbook ไม่สำเร็จ
Earned Value generation failed. => สร้าง Earned Value ไม่สำเร็จ
Live Payment uses the saved Payment Input and progress data. => Live Payment ใช้ข้อมูล Payment Input และผลงานที่บันทึกไว้
Payment rebuild replaces Payment only. Progress, monthly and dashboard sheets are preserved. => อัปเดตเฉพาะ Payment โดยรักษาชีต Progress รายเดือน และ Dashboard
Select a workbook first. => กรุณาเลือก Workbook ก่อน
Select an EV-ready workbook first. => กรุณาเลือก Workbook ที่พร้อมสร้าง EV ก่อน
Rebuild Workbook => อัปเดต Workbook
Rebuild generated views from the workbook itself. No Progress Studio session, BOQ file, or XML source is required. => อัปเดตมุมมองจาก Workbook โดยตรง ไม่ต้องใช้ session, ไฟล์ BOQ หรือ XML ต้นฉบับ
Select the saved Excel workbook containing your latest progress edits. => เลือกไฟล์ Excel ที่บันทึกผลงานล่าสุดแล้ว
Output Mode => รูปแบบผลลัพธ์
Snapshot Workbook => Workbook แบบ Snapshot
Refresh generated views from saved workbook edits. => อัปเดตมุมมองจากข้อมูลที่แก้ไขและบันทึกไว้
Live Workbook => Workbook แบบ Live
Keep supported views linked to Excel inputs. Use F9 / Save to recalculate. => มุมมองที่รองรับเชื่อมกับข้อมูล Excel ใช้ F9 / Save เพื่อคำนวณ
Target => งานที่จะอัปเดต
Refresh progress, monthly reporting and Dashboard views. => อัปเดตความก้าวหน้า รายงานรายเดือน และ Dashboard
Refresh Payment lines from main and Payment Input. => อัปเดตเส้น Payment จาก main และ Payment Input
Earned Value — Live => Earned Value — Live
Generate / Refresh the Earned Value view from current main progress and the selected monetary source. BAC changes require EV refresh. => สร้างหรืออัปเดต Earned Value จากผลงานใน main และแหล่งมูลค่าที่เลือก หาก BAC เปลี่ยนต้องอัปเดต EV
Build => สร้าง / อัปเดต
Check the selected monetary source and resolve the reported input issues before generating EV. => ตรวจแหล่งมูลค่าที่เลือกและแก้ข้อมูลตามที่แจ้งก่อนสร้าง EV
Generate / Refresh EV => สร้าง / อัปเดต EV
Activity Amount => Activity Amount
BOQ Mapping => BOQ Mapping
Equal => เท่ากัน
Duration => ตามระยะเวลา
Amount => ตามมูลค่า
'''.strip().splitlines())

PAIRS.update({
    "No matching activities": "ไม่พบกิจกรรมที่ตรงกับการค้นหา",
    "No matching BOQ items": "ไม่พบรายการ BOQ ที่ตรงกับการค้นหา",
    "Not mapped": "ยังไม่จับคู่", "Allocated": "จัดสรรแล้ว", "Remaining %": "% คงเหลือ",
    "Mapped To": "จับคู่กับ", "Activity ID / WBS": "Activity ID / WBS",
    "Activity": "กิจกรรม", "XML value (unrounded)": "ค่า XML (ไม่ปัดเศษ)",
    "Weight (display)": "น้ำหนัก (แสดงผล)", "Validation": "ผลตรวจสอบ",
    "Generating workbook...": "กำลังสร้าง Workbook...", "Workbook exported": "ส่งออก Workbook แล้ว",
    "Generating Workbook": "กำลังสร้าง Workbook",
    "Progress Studio is rebuilding the workbook from the working tree.": "Progress Studio กำลังสร้าง Workbook จากโครงสร้างงานปัจจุบัน",
    "Starting...": "กำลังเริ่มต้น...",
    "Read working schedule": "อ่านแผนงานปัจจุบัน",
    "Build main schedule": "สร้างแผนงานหลัก",
    "Build timescale": "สร้างช่วงเวลา",
    "Apply amount mapping": "กำหนดมูลค่าที่จัดสรร",
    "Build progress sheets + Dashboard": "สร้างชีตความก้าวหน้าและ Dashboard",
    "Finalize workbook": "จัดเตรียม Workbook ขั้นสุดท้าย",
    "Choose a field to preview.": "เลือกฟิลด์เพื่อดูตัวอย่าง",
    "No declared numeric activity/task custom fields were found in this XML.": "ไม่พบฟิลด์ตัวเลขแบบกำหนดเองของกิจกรรมใน XML นี้",
    "Import Schedule XML": "นำเข้าแผนงาน XML",
    "Prepare plan / actual schedule": "เตรียมแผนและผลงานจริง",
    "Build weekly timescale": "สร้างช่วงเวลารายสัปดาห์",
    "Build amount mapping": "กำหนดน้ำหนักกิจกรรม",
    "Build progress workbook": "สร้าง Progress Workbook",
    "Generate plan distribution": "กระจายแผนงาน",
    "Build OKD sheets": "สร้างชีตประกอบ",
    "Build monthly main view": "สร้างมุมมองรายเดือน",
    "Output created:\n{value1}": "สร้างผลลัพธ์แล้ว:\n{value1}",
    "Step {value1}/{value2}: {value3}": "ขั้นตอน {value1}/{value2}: {value3}",
    "Progress Studio {value1}": "Progress Studio {value1}",
    "Completed step {value1} of {value2}": "เสร็จขั้นตอน {value1} จาก {value2}",
    "Rows {value1}-{value2} of {value3} | Page {value4}/{value5}": "แถว {value1}-{value2} จาก {value3} | หน้า {value4}/{value5}",
    "Selected {value1:,} items | {value2:,.2f}": "เลือก {value1:,} รายการ | {value2:,.2f}",
    "{value1} {value2:,} BOQ items for {value3}?\nSelected amount: {value4:,.2f}": "{value1} BOQ {value2:,} รายการสำหรับ {value3}?\nมูลค่าที่เลือก: {value4:,.2f}",
    "Mapped {value1:,.2f} / {value2:,.2f} | Remaining {value3:,.2f} | Items {value4}/{value5}": "จับคู่ {value1:,.2f} / {value2:,.2f} | คงเหลือ {value3:,.2f} | รายการ {value4}/{value5}",
    "Project: {value1} (saved)": "โครงการ: {value1} (บันทึกแล้ว)",
    "Delete {value1}?\n\nChild nodes will also be removed from the working tree. The source workbook will not be overwritten.": "ลบ {value1}?\n\nรายการย่อยจะถูกลบจากโครงสร้างงานด้วย แต่ไม่เขียนทับ Workbook ต้นฉบับ",
    "Project saved:\n{value1}": "บันทึกโครงการแล้ว:\n{value1}",
    "Project: {value1} (loaded)": "โครงการ: {value1} (โหลดแล้ว)",
    "Move {value1} — {value2} under:": "ย้าย {value1} — {value2} ไปใต้:",
    "Auto-save failed:\n{value1}": "บันทึกอัตโนมัติไม่สำเร็จ:\n{value1}",
    "{value1} workbook cannot be verified at its saved location.\n\nSaved file: {value2}\n\nThis legacy project does not contain an embedded workbook copy.\nBrowse for the moved or renamed workbook?": "ตรวจสอบ Workbook {value1} ที่เดิมไม่ได้\n\nไฟล์เดิม: {value2}\n\nโครงการรุ่นเก่านี้ไม่มีสำเนา Workbook ฝังอยู่\nต้องการเลือกไฟล์ที่ย้ายหรือเปลี่ยนชื่อหรือไม่?",
    "Mapped workbook created:\n{value1}\n\nAmount rows updated: {value2}\nMapping rows written: {value3}\n\n{value4}{value5}": "สร้าง Workbook ที่จับคู่แล้ว:\n{value1}\n\nแถว Amount ที่อัปเดต: {value2}\nแถว Mapping ที่บันทึก: {value3}\n\n{value4}{value5}",
    "Ready • main found • {value1:,} activities": "พร้อม • พบ main • {value1:,} กิจกรรม",
    "Created {value1} • {value2} periods • {value3:,} activities • {value4:,} preserved": "สร้าง {value1} • {value2} งวด • {value3:,} กิจกรรม • รักษาข้อมูลเดิม {value4:,}",
    "Created {value1} • {value2:,} derived activities • {value3:,} eligible source activities • {value4:,} skipped": "สร้าง {value1} • สรุป {value2:,} กิจกรรม • ต้นทางที่ใช้ {value3:,} กิจกรรม • ข้าม {value4:,}",
    "Existing Payment Input • {value1} periods • {value2:,} requirements.": "พบ Payment Input • {value1} งวด • {value2:,} รายการ",
    "Not ready — {value1}": "ยังไม่พร้อม — {value1}",
    "Default {value1} periods from {value2:%d-%b-%y} to {value3:%d-%b-%y}.": "ค่าเริ่มต้น {value1} งวด ตั้งแต่ {value2:%d-%b-%y} ถึง {value3:%d-%b-%y}",
    "Could not open workbook:\n{value1}": "เปิด Workbook ไม่ได้:\n{value1}",
    "Ready • main found • {value1:,} activities • {value2}": "พร้อม • พบ main • {value1:,} กิจกรรม • {value2}",
    "Ready • {value1:,} activities • {value2}": "พร้อม • {value1:,} กิจกรรม • {value2}",
    "Progress rebuild will replace {value1} generated sheets ({value2} currently present). main is preserved.": "อัปเดตความก้าวหน้าจะแทนที่ชีตผลลัพธ์ {value1} ชีต (ปัจจุบันมี {value2}) โดยรักษา main",
    "Name": "ชื่อ", "Description": "รายละเอียด", "Status": "สถานะ", "Unit": "หน่วย",
    "Quantity": "ปริมาณ", "Rate": "ราคา", "Remaining": "คงเหลือ", "Mapped": "จับคู่แล้ว",
    "Activity ID": "Activity ID", "WBS": "WBS", "All": "ทั้งหมด",
    "Wait for the current operation to finish before closing.": "กรุณารอให้งานปัจจุบันเสร็จก่อนปิดโปรแกรม",
})
