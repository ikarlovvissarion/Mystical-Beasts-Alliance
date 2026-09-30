# MYSTICAL BEASTS ALLIANCE — WORK LOG writing setup

## สิ่งที่เพิ่ม
- หน้า WORK LOG มีฟอร์มกรอกข้อมูลจริง
- Target ID ดึงจาก BEASTS + PLANTS และ Target Name เติมอัตโนมัติ
- ผู้ลงบันทึกดึงจาก MEMBER
- เบิกคลังชมรมดึงจาก INVENTORY และเติมหน่วยอัตโนมัติ
- เก็บผลผลิตดึงรายการจาก INVENTORY และเพิ่มจำนวนเข้าคลัง
- เบิกจะหักจำนวนจาก INVENTORY
- ทุกการเปลี่ยนคลังสร้างรายการใน INVENTORY LOG และผูกด้วย Work ID
- Work ID สร้างอัตโนมัติ เช่น W0001, W0002...
- Transaction ID สร้างอัตโนมัติ เช่น T00001, T00002...

## ตั้งค่า Apps Script
1. เปิด Google Sheet ปัจจุบัน
2. Extensions > Apps Script
3. สร้าง/แทนที่ไฟล์ Code.gs ด้วยเนื้อหาใน `Code.gs`
4. Deploy > New deployment > Web app
5. Execute as: Me
6. Who has access: Anyone
7. กด Deploy และคัดลอก Web app URL
8. เปิด `index.html` แล้วแทนค่า `WORKLOG_API_URL` จาก `PASTE_APPS_SCRIPT_WEB_APP_URL_HERE` เป็น URL ที่ได้
9. อัปโหลด index.html + styles.css + Code.gs (Apps Script แยก deploy) ไปยัง GitHub Pages ตามปกติ

## หมายเหตุ
- การเบิกจะไม่ยอมให้ยอดติดลบ
- การรับผลผลิตต้องเลือกรายการใน INVENTORY ก่อน หากเป็นผลผลิตชนิดใหม่ ให้เพิ่มรายการนั้นใน INVENTORY master ก่อน
- รายการ "วันผู้วิเศษ" ใน UI ตอนนี้ใช้ชื่อวันจันทร์-อาทิตย์เป็นค่าเริ่มต้น และแก้ได้ใน index.html/Code.gs ให้ตรงกับระบบของชมรม
