# Mystical Beasts Alliance — Current Web App

เว็บแอปแบบ Static สำหรับ MYSTICAL BEASTS ALLIANCE โดยคงโครงสร้างและหน้าตาจากเค้าโครงเดิม

## Current Google Sheet
Spreadsheet ID:
`15HEODMU3gdVnkJfdD8CUzls51zcTyQWs`

เว็บจะอ่านแท็บ:
- `BEASTS`
- `PLANTS`
- `WORK LOG`
- `INVENTORY`

และตรวจข้อมูลใหม่ทุก 15 วินาที

## BEASTS
โครงสร้างปัจจุบัน:
`ID | Name | Species | Category | Breed | Classification | Sex | Age | Color | Date of Birth | Origin/Native Habitat | Other Characteristics | Owner | Image`

Category:
- A — สัตว์เลี้ยงลูกด้วยนม / MAMMALS
- B — สัตว์ปีกและแมลง / BIRDS & INSECTS
- C — สัตว์น้ำ สัตว์ครึ่งบกครึ่งน้ำ และสัตว์เลื้อยคลาน / AQUATIC, AMPHIBIANS & REPTILES

เว็บจัดกลุ่มตาม Category อัตโนมัติ

## Important
การอ่าน Google Sheets ผ่าน gviz ต้องให้ Spreadsheet สามารถเข้าถึงได้จากเว็บตามการตั้งค่าการแชร์ของ Google Sheets
หากอ่านไม่ได้ เว็บจะใช้ข้อมูล BEASTS สำรอง 7 รายการจากไฟล์ล่าสุดที่ส่งมา

## Deployment
อัปโหลด `index.html`, `styles.css` และไฟล์ที่ต้องการขึ้น GitHub Pages ได้เลย
