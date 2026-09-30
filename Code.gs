const SPREADSHEET_ID = '1pjxixoLuVNNbbQMaUMoPj2_nF1K0Ebibg3-MyrikeO0';

const WIZARD_DAYS = [
  'วันที่ 1 ผู้วิเศษ | 02.00 - 05.59 น.',
  'วันที่ 2 ผู้วิเศษ | 06.00 - 09.59 น.',
  'วันที่ 3 ผู้วิเศษ | 10.00 - 13.59 น.',
  'วันที่ 4 ผู้วิเศษ | 14.00 - 17.59 น.',
  'วันที่ 5 ผู้วิเศษ | 18.00 - 21.59 น.',
  'วันที่ 6 ผู้วิเศษ | 22.00 - 01.59 น.'
];

const WORK_TYPES = {
  'งานทั่วไป': [
    'การทำความสะอาด',
    'การให้อาหารสัตว์วิเศษ',
    'การรดน้ำพืชผัก'
  ],
  'เก็บผลผลิต': [
    'การเก็บผลผลิตพืชผัก',
    'การรับผลผลิตจากสัตว์วิเศษ'
  ]
};

function doGet() {
  return json_({ ok: true, service: 'MYSTICAL BEASTS ALLIANCE WORK LOG API' });
}

function doPost(e) {
  try {
    const payloadText = e && e.parameter && e.parameter.payload;
    if (!payloadText) throw new Error('ไม่พบ payload');
    const data = JSON.parse(payloadText);
    if (data.action === 'cancelWork') {
      const result = cancelWorkLog_(data);
      return json_(result);
    }
    const result = saveWorkLog_(data);
    return json_(result);
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

function saveWorkLog_(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const workSheet = getOrCreateSheet_(ss, 'WORK LOG');
    const inventorySheet = getOrCreateSheet_(ss, 'INVENTORY');
    const inventoryLogSheet = getOrCreateSheet_(ss, 'INVENTORY LOG');

    validate_(data);

    const workHeaders = ensureHeaders_(workSheet, [
      'Category', 'Work ID', 'วันที่ (มักเกิ้ล)', 'วันผู้วิเศษ', 'เวลา (IC)',
      'ประเภทการทำงาน', 'Target ID', 'Target Name', 'ผู้ลงบันทึก',
      'เบิกคลังชมรม', 'จำนวนที่เบิก', 'หน่วยที่เบิก', 'จำนวนที่ได้รับ', 'หน่วยที่ได้รับ', 'รูปการทำงาน'
    ]);
    const inventoryHeaders = ensureHeaders_(inventorySheet, ['Category', 'ID', 'รายการ', 'จำนวน', 'หน่วย']);
    const inventoryLogHeaders = ensureHeaders_(inventoryLogSheet, [
      'Transaction ID', 'วันที่ (มักเกิ้ล)', 'เวลา (IC)', 'ประเภท',
      'Item ID', 'รายการ', 'จำนวน', 'หน่วย', 'Work ID', 'ผู้ลงบันทึก', 'หมายเหตุ'
    ]);

    const workId = nextId_(workSheet, 'Work ID', 'W', 4);
    const now = new Date();
    // ยึดวันที่มักเกิ้ลจากวันที่ปัจจุบัน ณ เวลาที่บันทึกจริง ไม่ใช้วันที่เก่าจากหน้าเว็บ
    const dateValue = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const withdrawals = normalizeItems_(data.withdrawals);
    const receives = normalizeItems_(data.receives);

    // ตรวจสอบรายการคลังทั้งหมดก่อนเขียน เพื่อไม่ให้เกิดรายการบางส่วนเมื่อมีข้อผิดพลาด
    const checkedWithdrawals = [];
    if (data.category === 'งานทั่วไป') {
      for (const item of withdrawals) {
        const found = findInventoryItem_(inventorySheet, inventoryHeaders, item.id, item.name);
        if (!found) throw new Error('ไม่พบรายการที่ต้องการเบิกใน INVENTORY: ' + (item.name || item.id));
        const qty = Number(item.qty);
        if (!(qty > 0)) throw new Error('จำนวนที่เบิกต้องมากกว่า 0');
        if (toNumber_(found.quantity) < qty) throw new Error('จำนวนในคลังไม่เพียงพอ: ' + found.name);
        checkedWithdrawals.push({ item, found, qty });
      }
    }

    const checkedReceives = [];
    if (data.category === 'เก็บผลผลิต') {
      if (!receives.length) throw new Error('กรุณาเพิ่มรายการผลผลิตที่รับเข้าอย่างน้อย 1 รายการ');
      for (const item of receives) {
        const found = findInventoryItem_(inventorySheet, inventoryHeaders, item.id, item.name);
        if (!found) throw new Error('ไม่พบรายการผลผลิตใน INVENTORY: ' + (item.name || item.id));
        const qty = Number(item.qty);
        if (!(qty > 0)) throw new Error('จำนวนที่รับเข้าต้องมากกว่า 0');
        checkedReceives.push({ item, found, qty });
      }
    }

    const workRow = blankRow_(workHeaders.length);
    setCell_(workRow, workHeaders, 'Category', data.category);
    setCell_(workRow, workHeaders, 'Work ID', workId);
    setCell_(workRow, workHeaders, 'วันที่ (มักเกิ้ล)', dateValue);
    setCell_(workRow, workHeaders, 'วันผู้วิเศษ', data.magicalDay);
    setCell_(workRow, workHeaders, 'เวลา (IC)', data.icTime);
    setCell_(workRow, workHeaders, 'ประเภทการทำงาน', data.workType);
    setCell_(workRow, workHeaders, 'Target ID', data.targetId || '');
    setCell_(workRow, workHeaders, 'Target Name', data.targetName || '');
    setCell_(workRow, workHeaders, 'ผู้ลงบันทึก', data.recorder);

    // คงคอลัมน์เดิมไว้เพื่อรองรับฐานข้อมูลเดิม โดยสรุปรายการหลายรายการเป็นข้อความใน WORK LOG
    if (checkedWithdrawals.length) {
      setCell_(workRow, workHeaders, 'เบิกคลังชมรม', checkedWithdrawals.map(x => x.found.name).join(' | '));
      setCell_(workRow, workHeaders, 'จำนวนที่เบิก', checkedWithdrawals.map(x => String(x.qty)).join(' | '));
      setCell_(workRow, workHeaders, 'หน่วยที่เบิก', checkedWithdrawals.map(x => x.found.unit || x.item.unit || '').join(' | '));
    }
    if (checkedReceives.length) {
      setCell_(workRow, workHeaders, 'จำนวนที่ได้รับ', checkedReceives.map(x => String(x.qty)).join(' | '));
      setCell_(workRow, workHeaders, 'หน่วยที่ได้รับ', checkedReceives.map(x => x.found.unit || x.item.unit || '').join(' | '));
    }

    // บันทึกรูปการทำงานลง Google Drive แล้วเก็บลิงก์รูปไว้ใน WORK LOG
    // ใช้ลิงก์ thumbnail ที่หน้าเว็บสามารถแสดงรูปได้โดยตรง
    if (data.photo && data.photo.data) {
      const photoUrl = saveWorkPhoto_(data.photo, workId);
      setCell_(workRow, workHeaders, 'รูปการทำงาน', photoUrl);
    }

    workSheet.appendRow(workRow);

    const transactions = [];

    for (const x of checkedWithdrawals) {
      const stock = toNumber_(x.found.quantity);
      const unit = x.found.unit || x.item.unit || '';
      inventorySheet.getRange(x.found.row, x.found.quantityCol).setValue(stock - x.qty);
      const txId = nextId_(inventoryLogSheet, 'Transaction ID', 'T', 5);
      appendInventoryLog_(inventoryLogSheet, inventoryLogHeaders, {
        txId, dateValue, time: data.icTime, type: 'เบิก', itemId: x.found.id,
        name: x.found.name, qty: x.qty, unit, workId, recorder: data.recorder,
        note: data.workType
      });
      transactions.push({ id: txId, type: 'เบิก', item: x.found.name, quantity: x.qty, unit });
    }

    for (const x of checkedReceives) {
      const stock = toNumber_(x.found.quantity);
      const unit = x.found.unit || x.item.unit || '';
      inventorySheet.getRange(x.found.row, x.found.quantityCol).setValue(stock + x.qty);
      const txId = nextId_(inventoryLogSheet, 'Transaction ID', 'T', 5);
      appendInventoryLog_(inventoryLogSheet, inventoryLogHeaders, {
        txId, dateValue, time: data.icTime, type: 'รับเข้า', itemId: x.found.id,
        name: x.found.name, qty: x.qty, unit, workId, recorder: data.recorder,
        note: data.workType
      });
      transactions.push({ id: txId, type: 'รับเข้า', item: x.found.name, quantity: x.qty, unit });
    }

    SpreadsheetApp.flush();
    return { ok: true, workId, transactions };
  } finally {
    lock.releaseLock();
  }
}


function saveWorkPhoto_(photo, workId) {
  const base64 = String(photo.data || '').trim();
  if (!base64) return '';

  const mimeType = String(photo.mimeType || 'image/jpeg').trim() || 'image/jpeg';
  if (!/^image\/(jpeg|png|webp)$/i.test(mimeType)) {
    throw new Error('รองรับเฉพาะ JPG, PNG หรือ WEBP');
  }

  const folderName = 'MYSTICAL BEASTS ALLIANCE - WORK PHOTOS';
  const folders = DriveApp.getFoldersByName(folderName);
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);

  const bytes = Utilities.base64Decode(base64);
  const safeName = String(photo.name || ('WORK-' + workId + '.jpg'))
    .replace(/[^a-zA-Z0-9._-]+/g, '_')
    .replace(/^_+|_+$/g, '') || ('WORK-' + workId + '.jpg');
  const blob = Utilities.newBlob(bytes, mimeType, workId + '-' + safeName);
  const file = folder.createFile(blob);

  // ทำให้รูปดูได้จากหน้าเว็บ GitHub Pages โดยไม่ต้องล็อกอิน Google
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (sharingError) {
    throw new Error('อัปโหลดรูปแล้ว แต่ตั้งสิทธิ์ดูรูปสาธารณะไม่สำเร็จ: ' + sharingError.message);
  }

  const fileId = file.getId();
  return 'https://drive.google.com/thumbnail?id=' + encodeURIComponent(fileId) + '&sz=w1200';
}

function validate_(data) {
  if (!data.category || !['งานทั่วไป', 'เก็บผลผลิต'].includes(data.category)) throw new Error('หมวดหมู่งานไม่ถูกต้อง');
  if (!data.workType || !WORK_TYPES[data.category].includes(data.workType)) throw new Error('ประเภทการทำงานไม่ถูกต้อง');
  if (data.category === 'งานทั่วไป' && data.workType !== 'การทำความสะอาด') {
    if (!data.targetId || !data.targetName) throw new Error('กรุณาเลือกสัตว์วิเศษ/พืชผัก');
  }
  if (data.category === 'เก็บผลผลิต' && (!data.targetId || !data.targetName)) throw new Error('กรุณาเลือกสัตว์วิเศษ/พืชผัก');
  if (!data.recorder) throw new Error('กรุณาเลือกผู้ลงบันทึก');
  if (!data.muggleDate) throw new Error('กรุณาเลือกวันที่');
  if (!WIZARD_DAYS.includes(data.magicalDay)) throw new Error('วันผู้วิเศษไม่ถูกต้อง');
  if (!data.icTime) throw new Error('กรุณาเลือกเวลา IC');
}

function normalizeItems_(items) {
  if (!Array.isArray(items)) return [];
  return items.map(x => ({
    id: String(x && x.id || '').trim(),
    name: String(x && x.name || '').trim(),
    qty: x && x.qty,
    unit: String(x && x.unit || '').trim()
  })).filter(x => x.id || x.name || x.qty);
}

function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}
function ensureHeaders_(sheet, expected) {
  const lastCol = Math.max(sheet.getLastColumn(), expected.length);
  const existing = sheet.getLastRow() ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String) : [];
  if (!sheet.getLastRow()) { sheet.getRange(1, 1, 1, expected.length).setValues([expected]); return expected; }
  const headers = existing.slice(0, Math.max(existing.length, expected.length));
  expected.forEach((h, i) => { if (!headers[i]) headers[i] = h; });
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  return headers;
}
function blankRow_(length) { return Array.from({ length }, () => ''); }
function setCell_(row, headers, name, value) { const i = headers.indexOf(name); if (i >= 0) row[i] = value; }
function findHeader_(headers, names) { for (const n of names) { const i = headers.indexOf(n); if (i >= 0) return i; } return -1; }

function findInventoryItem_(sheet, headers, id, name) {
  const idCol = findHeader_(headers, ['ID', 'Item ID']);
  const nameCol = findHeader_(headers, ['รายการ', 'Name', 'Item']);
  const qtyCol = findHeader_(headers, ['จำนวน', 'Quantity']);
  const unitCol = findHeader_(headers, ['หน่วย', 'Unit']);
  if (nameCol < 0 || qtyCol < 0) return null;
  const count = Math.max(sheet.getLastRow() - 1, 0);
  const values = count ? sheet.getRange(2, 1, count, headers.length).getValues() : [];
  for (let r = 0; r < values.length; r++) {
    const row = values[r];
    const rowId = idCol >= 0 ? String(row[idCol] || '').trim() : '';
    const rowName = String(row[nameCol] || '').trim();
    if ((id && rowId === String(id).trim()) || (name && rowName === String(name).trim())) {
      return { row: r + 2, id: rowId, name: rowName, quantity: row[qtyCol], unit: unitCol >= 0 ? String(row[unitCol] || '') : '', quantityCol: qtyCol + 1 };
    }
  }
  return null;
}
function appendInventoryLog_(sheet, headers, tx) {
  const row = blankRow_(headers.length);
  setCell_(row, headers, 'Transaction ID', tx.txId); setCell_(row, headers, 'วันที่ (มักเกิ้ล)', tx.dateValue);
  setCell_(row, headers, 'เวลา (IC)', tx.time); setCell_(row, headers, 'ประเภท', tx.type); setCell_(row, headers, 'Item ID', tx.itemId);
  setCell_(row, headers, 'รายการ', tx.name); setCell_(row, headers, 'จำนวน', tx.qty); setCell_(row, headers, 'หน่วย', tx.unit);
  setCell_(row, headers, 'Work ID', tx.workId); setCell_(row, headers, 'ผู้ลงบันทึก', tx.recorder); setCell_(row, headers, 'หมายเหตุ', tx.note || '');
  sheet.appendRow(row);
}
function nextId_(sheet, headerName, prefix, width) {
  const headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0].map(String);
  const col = headers.indexOf(headerName);
  if (col < 0 || sheet.getLastRow() < 2) return prefix + String(1).padStart(width, '0');
  const values = sheet.getRange(2, col + 1, sheet.getLastRow() - 1, 1).getValues().flat();
  let max = 0; values.forEach(v => { const m = String(v || '').match(/(\d+)$/); if (m) max = Math.max(max, Number(m[1])); });
  return prefix + String(max + 1).padStart(width, '0');
}
function toNumber_(v) { if (typeof v === 'number') return v; const n = Number(String(v || '').replace(/,/g, '').trim()); return Number.isFinite(n) ? n : 0; }
function json_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }

// ============================================================
// ADMIN: Google Account permission + Work Log cancellation
// ============================================================
const ADMIN_PASSWORD = 'Meduza2014';
function adminStatus_() { return {ok:true,isAdmin:false,auth:'password'}; }
function ensureColumn_(sheet, headers, name) { let i=headers.indexOf(name); if(i>=0)return i+1; const c=sheet.getLastColumn()+1; sheet.getRange(1,c).setValue(name); return c; }
function valueByHeader_(row, headers, name) { const i=headers.indexOf(name); return i>=0?row[i]:''; }
function cancelWorkLog_(data) {
  if (String(data.password || '') !== ADMIN_PASSWORD) throw new Error('Password ไม่ถูกต้อง');
  const email='Password Admin';
  const workId=String(data.workId||'').trim(), reason=String(data.reason||'').trim();
  if(!workId) throw new Error('ไม่พบ Work ID'); if(!reason) throw new Error('กรุณาระบุเหตุผลการยกเลิก');
  const lock=LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const ss=SpreadsheetApp.openById(SPREADSHEET_ID), ws=ss.getSheetByName('WORK LOG'), inv=ss.getSheetByName('INVENTORY'), il=ss.getSheetByName('INVENTORY LOG');
    if(!ws||!inv||!il) throw new Error('ไม่พบแท็บฐานข้อมูลที่จำเป็น');
    let headers=ws.getRange(1,1,1,Math.max(ws.getLastColumn(),1)).getValues()[0].map(String);
    const statusCol=ensureColumn_(ws,headers,'Status'), byCol=ensureColumn_(ws,headers,'Cancelled By'), atCol=ensureColumn_(ws,headers,'Cancelled At'), reasonCol=ensureColumn_(ws,headers,'Cancel Reason');
    headers=ws.getRange(1,1,1,ws.getLastColumn()).getValues()[0].map(String);
    const widCol=headers.indexOf('Work ID'), rows=ws.getLastRow()>1?ws.getRange(2,1,ws.getLastRow()-1,ws.getLastColumn()).getValues():[];
    let rowNo=-1,row=null; for(let i=0;i<rows.length;i++){if(String(rows[i][widCol]||'').trim()===workId){rowNo=i+2;row=rows[i];break;}}
    if(rowNo<0) throw new Error('ไม่พบ WORK ID: '+workId);
    if(String(row[statusCol-1]||'').toLowerCase()==='cancelled') throw new Error('รายการนี้ถูกยกเลิกไปแล้ว');
    const ih=inv.getRange(1,1,1,Math.max(inv.getLastColumn(),1)).getValues()[0].map(String), lh=il.getRange(1,1,1,Math.max(il.getLastColumn(),1)).getValues()[0].map(String);
    const logs=il.getLastRow()>1?il.getRange(2,1,il.getLastRow()-1,il.getLastColumn()).getValues():[], reverted=[];
    for(const lr of logs){
      if(String(valueByHeader_(lr,lh,'Work ID')||'').trim()!==workId) continue;
      const type=String(valueByHeader_(lr,lh,'ประเภท')||'').trim(), id=String(valueByHeader_(lr,lh,'Item ID')||'').trim(), name=String(valueByHeader_(lr,lh,'รายการ')||'').trim(), qty=toNumber_(valueByHeader_(lr,lh,'จำนวน'));
      if(!(qty>0))continue; const found=findInventoryItem_(inv,ih,id,name); if(!found)throw new Error('ไม่พบรายการใน INVENTORY: '+name);
      let newStock=toNumber_(found.quantity), rollback=''; if(type==='เบิก'){newStock+=qty;rollback='ปรับเพิ่ม';} else if(type==='รับเข้า'){if(newStock<qty)throw new Error('จำนวนในคลังไม่พอสำหรับย้อนผลผลิต: '+found.name);newStock-=qty;rollback='ปรับลด';} else continue;
      inv.getRange(found.row,found.quantityCol).setValue(newStock);
      const txId=nextId_(il,'Transaction ID','T',5); appendInventoryLog_(il,lh,{txId,dateValue:new Date(),time:Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'HH:mm'),type:rollback,itemId:found.id,name:found.name,qty:qty,unit:found.unit,workId:workId,recorder:email,note:'ยกเลิก '+workId+' | '+reason});
      reverted.push({item:found.name,quantity:qty,unit:found.unit,type:rollback});
    }
    // เก็บสำเนาไว้ใน DELETED WORK LOG ก่อนลบจริง เพื่อให้ ADMIN ยังดูประวัติย้อนหลังได้
    const deletedSheet=getOrCreateSheet_(ss,'DELETED WORK LOG');
    const deletedHeaders=ensureHeaders_(deletedSheet, headers.concat(['Deleted At','Deleted By','Delete Reason']));
    const deletedRow=blankRow_(deletedHeaders.length);
    headers.forEach((h,i)=>{ if(i<row.length) deletedRow[i]=row[i]; });
    setCell_(deletedRow, deletedHeaders, 'Deleted At', new Date());
    setCell_(deletedRow, deletedHeaders, 'Deleted By', email);
    setCell_(deletedRow, deletedHeaders, 'Delete Reason', reason);
    deletedSheet.appendRow(deletedRow);

    // ลบรายการ WORK LOG ออกจากฐานข้อมูลจริง หลังย้อนรายการคลังและเก็บประวัติแล้ว
    ws.deleteRow(rowNo); SpreadsheetApp.flush();
    return {ok:true,workId:workId,deleted:true,cancelledBy:email,reverted:reverted};
  } finally { lock.releaseLock(); }
}
