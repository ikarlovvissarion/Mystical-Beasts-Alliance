# แก้ WORK LOG ไม่บันทึก + ตรวจสิทธิ์ Google Drive

สาเหตุที่พบบ่อยที่สุดคือ Apps Script Web App ยังไม่ได้รับสิทธิ์ Drive/Sheets หรือ deployment ยังรัน Code.gs เวอร์ชันเก่า

## ทำตามนี้ตามลำดับ
1. เปิด Apps Script ที่เชื่อมกับ Google Sheet ID `1pjxixoLuVNNbbQMaUMoPj2_nF1K0Ebibg3-MyrikeO0`
2. แทนที่ `Code.gs` ด้วยไฟล์ใน ZIP นี้
3. กด Save
4. เลือกฟังก์ชัน `setupMBA` แล้วกด Run
5. Google จะถามสิทธิ์ ให้กดอนุญาต Google Sheets และ Google Drive
6. กลับไป Deploy > Manage deployments > Web app > Edit > New version > Deploy
7. ต้องตั้ง Execute as = Me (บัญชีเจ้าของ Apps Script)
8. Who has access ต้องเป็นค่าที่เว็บ GitHub Pages ใช้งานอยู่ เช่น Anyone
9. ใช้ URL /exec เดิมใน index.html

## ตรวจระบบ
เปิด URL:
`WEB_APP_URL?action=health`
ควรได้ JSON ที่มี:
- `spreadsheet: true`
- `driveFolder: true`
- `folderName: "MYSTICAL BEASTS ALLIANCE - WORK PHOTOS"`

ถ้า `driveFolder:false` แปลว่า Apps Script ยังไม่มีสิทธิ์เข้าถึงโฟลเดอร์
ถ้า `spreadsheet:false` แปลว่า Apps Script ยังไม่มีสิทธิ์/ใช้ Spreadsheet ID ผิด

หลัง health ผ่าน ให้ลองบันทึก WORK LOG พร้อมรูปใหม่
