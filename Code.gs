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

function doGet(e) {
  try {
    const action = e && e.parameter && e.parameter.action;
    if (action === 'health') return json_(healthCheck_());
    return json_({ ok: true, service: 'MYSTICAL BEASTS ALLIANCE WORK LOG API' });
  } catch (err) {
    return json_({ ok:false, error:String(err.message || err) });
  }
}

// รันฟังก์ชันนี้ 1 ครั้งจาก Apps Script Editor ด้วยบัญชีเจ้าของระบบ
// เพื่อบังคับให้ Google ขอสิทธิ์ Sheets + Drive ก่อนใช้งาน Web App
// ฟังก์ชันนี้จะตรวจทั้ง Spreadsheet และโฟลเดอร์รูป
function testMBASetup() {
  return healthCheck_();
}

function setupMBA() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const folder = DriveApp.getFolderById('1vUYwKD8Ka-sh8yo93qpr5AiKW9MHwDse');
  getOrCreateSheet_(ss, 'WORK LOG');
  getOrCreateSheet_(ss, 'INVENTORY');
  getOrCreateSheet_(ss, 'INVENTORY LOG');
  getOrCreateSheet_(ss, 'DELETED WORK LOG');
  getOrCreateSheet_(ss, 'DELETE REQUESTS');
  try { folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  return {ok:true, spreadsheet:ss.getName(), folder:folder.getName(), folderId:folder.getId()};
}

function healthCheck_() {
  const result = {ok:true, spreadsheet:false, driveFolder:false, folderName:'', folderId:'1vUYwKD8Ka-sh8yo93qpr5AiKW9MHwDse'};
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    result.spreadsheet = true;
    result.spreadsheetName = ss.getName();
    result.workLog = !!ss.getSheetByName('WORK LOG');
    result.inventory = !!ss.getSheetByName('INVENTORY');
  } catch (e) { result.ok=false; result.spreadsheetError=String(e.message||e); }
  try {
    const folder = DriveApp.getFolderById(result.folderId);
    result.driveFolder = true;
    result.folderName = folder.getName();
  } catch (e) { result.ok=false; result.driveError=String(e.message||e); }
  return result;
}

function doPost(e) {
  try {
    // รองรับทั้ง payload แบบ form parameter และ raw JSON body
    let payloadText = e && e.parameter && e.parameter.payload;
    if (!payloadText && e && e.postData && e.postData.contents) {
      payloadText = e.postData.contents;
    }
    if (!payloadText) throw new Error('ไม่พบ payload');

    const data = typeof payloadText === 'string'
      ? JSON.parse(payloadText)
      : payloadText;
    if (data.action === 'testPhoto') {
      return json_({ok:true, service:'photo-ready', spreadsheetId:SPREADSHEET_ID});
    }
    if (data.action === 'cancelWork') {
      const result = cancelWorkLog_(data);
      return json_(result);
    }
    if (data.action === 'requestDelete') return json_(requestDeleteWork_(data));
    if (data.action === 'approveDeleteRequest') return json_(approveDeleteRequest_(data));
    if (data.action === 'rejectDeleteRequest') return json_(rejectDeleteRequest_(data));
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
      'เบิกคลังชมรม', 'จำนวนที่เบิก', 'หน่วยที่เบิก', 'คืนคลังชมรม', 'จำนวนที่คืน', 'หน่วยที่คืน', 'จำนวนที่ได้รับ', 'หน่วยที่ได้รับ', 'รูปการทำงาน'
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
    const returns = normalizeItems_(data.returns);
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

    const checkedReturns = [];
    if (returns.length) {
      for (const item of returns) {
        const found = findInventoryItem_(inventorySheet, inventoryHeaders, item.id, item.name);
        if (!found) throw new Error('ไม่พบรายการที่ต้องการคืนใน INVENTORY: ' + (item.name || item.id));
        const qty = Number(item.qty);
        if (!(qty > 0)) throw new Error('จำนวนที่คืนต้องมากกว่า 0');
        const returnable = getOutstandingWithdrawn_(inventoryLogSheet, inventoryLogHeaders, found.id, found.name);
        if (qty > returnable) throw new Error('คืนเกินจำนวนที่เคยเบิก: ' + found.name + ' (คืนได้สูงสุด ' + returnable + ' ' + (found.unit || item.unit || '') + ')');
        checkedReturns.push({ item, found, qty });
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
    if (checkedReturns.length) {
      setCell_(workRow, workHeaders, 'คืนคลังชมรม', checkedReturns.map(x => x.found.name).join(' | '));
      setCell_(workRow, workHeaders, 'จำนวนที่คืน', checkedReturns.map(x => String(x.qty)).join(' | '));
      setCell_(workRow, workHeaders, 'หน่วยที่คืน', checkedReturns.map(x => x.found.unit || x.item.unit || '').join(' | '));
    }
    if (checkedReceives.length) {
      setCell_(workRow, workHeaders, 'จำนวนที่ได้รับ', checkedReceives.map(x => String(x.qty)).join(' | '));
      setCell_(workRow, workHeaders, 'หน่วยที่ได้รับ', checkedReceives.map(x => x.found.unit || x.item.unit || '').join(' | '));
    }

    // ต้องอัปโหลดรูปให้สำเร็จก่อน จึงจะบันทึก WORK LOG
    // เพื่อให้ทุก WORK LOG ที่เกิดขึ้นมีรูปและหน้าเว็บสามารถแสดงรูปได้แน่นอน
    const photoUrl = saveWorkPhoto_(data.photo, workId);
    if (!photoUrl) throw new Error('อัปโหลดรูปการทำงานไม่สำเร็จ');
    setCell_(workRow, workHeaders, 'รูปการทำงาน', photoUrl);

    // บันทึก WORK LOG หลังจากรูปพร้อมแล้ว
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

    for (const x of checkedReturns) {
      const stock = toNumber_(x.found.quantity);
      const unit = x.found.unit || x.item.unit || '';
      inventorySheet.getRange(x.found.row, x.found.quantityCol).setValue(stock + x.qty);
      const txId = nextId_(inventoryLogSheet, 'Transaction ID', 'T', 5);
      appendInventoryLog_(inventoryLogSheet, inventoryLogHeaders, {
        txId, dateValue, time: data.icTime, type: 'คืน', itemId: x.found.id,
        name: x.found.name, qty: x.qty, unit, workId, recorder: data.recorder,
        note: data.workType
      });
      transactions.push({ id: txId, type: 'คืน', item: x.found.name, quantity: x.qty, unit });
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
    return { ok: true, workId, transactions, photoUrl };
  } finally {
    lock.releaseLock();
  }
}


function saveWorkPhoto_(photo, workId) {
  photo = photo || {};

  // รองรับทั้ง {data: "..."} และ {base64: "..."} จากหน้าเว็บ
  let base64 = String(photo.data || photo.base64 || '').trim();
  if (!base64) return '';

  const mimeType = String(photo.mimeType || photo.type || 'image/jpeg').trim() || 'image/jpeg';
  if (!/^image\/(jpeg|jpg|png|webp)$/i.test(mimeType)) {
    throw new Error('รองรับเฉพาะ JPG, PNG หรือ WEBP');
  }

  // รองรับ Data URL เช่น data:image/jpeg;base64,xxxx
  if (base64.indexOf(',') >= 0 && /^data:/i.test(base64)) {
    base64 = base64.substring(base64.indexOf(',') + 1);
  }
  base64 = base64.replace(/\s/g, '');

  if (base64.length < 100) {
    throw new Error('ข้อมูลรูปภาพไม่สมบูรณ์');
  }
  if (base64.length > 12 * 1024 * 1024) {
    throw new Error('รูปภาพมีขนาดใหญ่เกินไป กรุณาเลือกรูปที่เล็กลง');
  }

  const WORK_PHOTOS_FOLDER_ID = '1vUYwKD8Ka-sh8yo93qpr5AiKW9MHwDse';

  let folder;
  try {
    folder = DriveApp.getFolderById(WORK_PHOTOS_FOLDER_ID);
  } catch (e) {
    throw new Error(
      'Apps Script เข้าถึงโฟลเดอร์ Google Drive ไม่ได้: ' +
      String(e.message || e)
    );
  }

  try {
    folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (sharingError) {
    console.warn('Folder sharing skipped: ' + sharingError.message);
  }

  let bytes;
  try {
    bytes = Utilities.base64Decode(base64);
  } catch (e) {
    throw new Error('ถอดรหัสรูปภาพไม่สำเร็จ กรุณาลองเลือกรูปใหม่');
  }

  const extension =
    /png/i.test(mimeType) ? 'png' :
    /webp/i.test(mimeType) ? 'webp' : 'jpg';

  const originalName = String(
    photo.name || ('WORK-' + workId + '.' + extension)
  ).trim();

  const safeName = originalName
    .replace(/[^a-zA-Z0-9._-]+/g, '_')
    .replace(/^_+|_+$/g, '') ||
    ('WORK-' + workId + '.' + extension);

  const blob = Utilities.newBlob(
    bytes,
    mimeType === 'image/jpg' ? 'image/jpeg' : mimeType,
    workId + '-' + safeName
  );

  let file;
  try {
    file = folder.createFile(blob);
  } catch (e) {
    throw new Error(
      'สร้างไฟล์รูปใน Google Drive ไม่สำเร็จ: ' +
      String(e.message || e)
    );
  }

  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (sharingError) {
    console.warn('File sharing skipped: ' + sharingError.message);
  }

  const fileId = file.getId();

  return 'https://drive.google.com/thumbnail?id=' +
    encodeURIComponent(fileId) + '&sz=w1200';
}

function validate_(data) {
  if (!data.category || !['งานทั่วไป', 'เก็บผลผลิต'].includes(data.category)) throw new Error('หมวดหมู่งานไม่ถูกต้อง');
  if (!data.workType || !WORK_TYPES[data.category].includes(data.workType)) throw new Error('ประเภทการทำงานไม่ถูกต้อง');
  if (data.category === 'งานทั่วไป' && data.workType !== 'การทำความสะอาด') {
    if (!data.targetId || !data.targetName) throw new Error('กรุณาเลือกสัตว์วิเศษ/พืชผัก');
  }
  if (data.category === 'เก็บผลผลิต' && (!data.targetId || !data.targetName)) throw new Error('กรุณาเลือกสัตว์วิเศษ/พืชผัก');
  if (!data.recorder) throw new Error('กรุณาเลือกผู้ลงบันทึก');
  if (!data.photo || !data.photo.data) throw new Error('กรุณาแนบรูปการทำงาน');
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
  // รักษาคอลัมน์เดิมทั้งหมด และเติมคอลัมน์ที่ระบบต้องใช้ซึ่งยังไม่มีจริง
  // สำคัญ: ถ้า WORK LOG มีหัวตารางเดิมครบอยู่แล้ว แต่ไม่มี "รูปการทำงาน"
  // ต้องเพิ่มหัวข้อนี้ต่อท้าย ไม่เช่นนั้น setCell_ จะหาไม่เจอและ URL จะหายไป
  if (!sheet.getLastRow()) {
    sheet.getRange(1, 1, 1, expected.length).setValues([expected]);
    return expected.slice();
  }

  const currentLastCol = Math.max(sheet.getLastColumn(), 1);
  const existing = sheet.getRange(1, 1, 1, currentLastCol).getValues()[0].map(v => String(v || '').trim());
  const headers = existing.slice();

  // เติมช่องว่างที่อยู่ในตำแหน่งเดิมก่อน เพื่อไม่ย้าย schema เดิม
  expected.forEach((h, i) => {
    if (!headers[i]) headers[i] = h;
  });

  // ถ้าหัวข้อที่ต้องการยังไม่มีทุกตำแหน่ง ให้ append ต่อท้าย
  for (const h of expected) {
    if (headers.indexOf(h) < 0) headers.push(h);
  }

  if (headers.length > currentLastCol) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
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
function getOutstandingWithdrawn_(sheet, headers, itemId, itemName) {
  const count = Math.max(sheet.getLastRow() - 1, 0);
  if (!count) return 0;
  const values = sheet.getRange(2, 1, count, headers.length).getValues();
  let withdrawn = 0, returned = 0;
  for (const row of values) {
    const id = String(valueByHeader_(row, headers, 'Item ID') || '').trim();
    const name = String(valueByHeader_(row, headers, 'รายการ') || '').trim();
    if (!((itemId && id === String(itemId).trim()) || (itemName && name === String(itemName).trim()))) continue;
    const type = String(valueByHeader_(row, headers, 'ประเภท') || '').trim();
    const qty = toNumber_(valueByHeader_(row, headers, 'จำนวน'));
    if (type === 'เบิก') withdrawn += qty;
    if (type === 'คืน') returned += qty;
    if (type === 'ปรับเพิ่ม' && String(valueByHeader_(row, headers, 'หมายเหตุ') || '').includes('ย้อนรายการ: เบิก')) withdrawn -= qty;
    if (type === 'ปรับลด' && String(valueByHeader_(row, headers, 'หมายเหตุ') || '').includes('ย้อนรายการ: คืน')) returned -= qty;
  }
  return Math.max(0, withdrawn - returned);
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
// ADMIN: Password permission + Work Log cancellation
// ============================================================
const ADMIN_PASSWORD = 'Meduza2014';
function adminStatus_() { return {ok:true,isAdmin:false,auth:'password'}; }
function ensureColumn_(sheet, headers, name) { let i=headers.indexOf(name); if(i>=0)return i+1; const c=sheet.getLastColumn()+1; sheet.getRange(1,c).setValue(name); return c; }
function valueByHeader_(row, headers, name) { const i=headers.indexOf(name); return i>=0?row[i]:''; }
function requestDeleteWork_(data) {
  const workId=String(data.workId||'').trim(), requester=String(data.requester||'').trim(), reason=String(data.reason||'').trim();
  if(!workId) throw new Error('ไม่พบ Work ID');
  if(!requester) throw new Error('กรุณาระบุชื่อผู้ขอลบ');
  if(!reason) throw new Error('กรุณาระบุเหตุผลการขอลบ');
  const lock=LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const ss=SpreadsheetApp.openById(SPREADSHEET_ID), ws=ss.getSheetByName('WORK LOG');
    const rs=getOrCreateSheet_(ss,'DELETE REQUESTS');
    const rh=ensureHeaders_(rs,['Request ID','Requested At','Work ID','ผู้ขอลบ','เหตุผล','Status','Reviewed At','Reviewed By','Admin Reason']);
    const wh=ws.getRange(1,1,1,Math.max(ws.getLastColumn(),1)).getValues()[0].map(String);
    const rows=ws.getLastRow()>1?ws.getRange(2,1,ws.getLastRow()-1,ws.getLastColumn()).getValues():[];
    const widCol=wh.indexOf('Work ID'), statusCol=wh.indexOf('Status');
    const row=rows.find(r=>String(r[widCol]||'').trim()===workId);
    if(!row) throw new Error('ไม่พบ WORK ID: '+workId);
    if(statusCol>=0 && String(row[statusCol]||'').toLowerCase()==='cancelled') throw new Error('รายการนี้ถูกยกเลิกแล้ว');
    const reqRows=rs.getLastRow()>1?rs.getRange(2,1,rs.getLastRow()-1,rs.getLastColumn()).getValues():[];
    const existing=reqRows.find(r=>String(valueByHeader_(r,rh,'Work ID')||'').trim()===workId && String(valueByHeader_(r,rh,'Status')||'').trim()==='PENDING');
    if(existing) throw new Error('รายการนี้มีคำขอลบที่รอแอดมินอยู่แล้ว');
    const reqId=nextId_(rs,'Request ID','D',5), out=blankRow_(rh.length);
    setCell_(out,rh,'Request ID',reqId); setCell_(out,rh,'Requested At',new Date()); setCell_(out,rh,'Work ID',workId); setCell_(out,rh,'ผู้ขอลบ',requester); setCell_(out,rh,'เหตุผล',reason); setCell_(out,rh,'Status','PENDING');
    rs.appendRow(out); SpreadsheetApp.flush(); return {ok:true,requestId:reqId};
  } finally { lock.releaseLock(); }
}
function approveDeleteRequest_(data) {
  if(String(data.password||'')!==ADMIN_PASSWORD) throw new Error('Password ไม่ถูกต้อง');
  const requestId=String(data.requestId||'').trim(), reason=String(data.reason||'').trim();
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID), rs=ss.getSheetByName('DELETE REQUESTS');
  if(!rs) throw new Error('ไม่พบแท็บ DELETE REQUESTS');
  const rh=rs.getRange(1,1,1,Math.max(rs.getLastColumn(),1)).getValues()[0].map(String);
  const rows=rs.getLastRow()>1?rs.getRange(2,1,rs.getLastRow()-1,rs.getLastColumn()).getValues():[];
  const idCol=rh.indexOf('Request ID'), statusCol=rh.indexOf('Status');
  let rowNo=-1,row=null; for(let i=0;i<rows.length;i++){if(String(rows[i][idCol]||'').trim()===requestId){rowNo=i+2;row=rows[i];break;}}
  if(rowNo<0) throw new Error('ไม่พบคำขอ: '+requestId);
  if(String(row[statusCol]||'')!=='PENDING') throw new Error('คำขอนี้ถูกดำเนินการแล้ว');
  const workId=String(valueByHeader_(row,rh,'Work ID')||'').trim();
  const result=cancelWorkLog_({workId,password:ADMIN_PASSWORD,reason:reason||String(valueByHeader_(row,rh,'เหตุผล')||'ลบตามคำขอ'),skipRequest:true,deleteSource:'DELETE REQUEST',requestId});
  setCell_(row,rh,'Status','APPROVED'); setCell_(row,rh,'Reviewed At',new Date()); setCell_(row,rh,'Reviewed By','Admin'); setCell_(row,rh,'Admin Reason',reason||'อนุมัติคำขอลบ');
  rs.getRange(rowNo,1,1,rh.length).setValues([row]); SpreadsheetApp.flush(); return {ok:true,requestId,workId,result};
}
function rejectDeleteRequest_(data) {
  if(String(data.password||'')!==ADMIN_PASSWORD) throw new Error('Password ไม่ถูกต้อง');
  const requestId=String(data.requestId||'').trim(), reason=String(data.reason||'').trim()||'ไม่อนุมัติ';
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID), rs=ss.getSheetByName('DELETE REQUESTS');
  if(!rs) throw new Error('ไม่พบแท็บ DELETE REQUESTS');
  const rh=rs.getRange(1,1,1,Math.max(rs.getLastColumn(),1)).getValues()[0].map(String), rows=rs.getLastRow()>1?rs.getRange(2,1,rs.getLastRow()-1,rs.getLastColumn()).getValues():[];
  const idCol=rh.indexOf('Request ID'), statusCol=rh.indexOf('Status'); let rowNo=-1,row=null;
  for(let i=0;i<rows.length;i++){if(String(rows[i][idCol]||'').trim()===requestId){rowNo=i+2;row=rows[i];break;}}
  if(rowNo<0) throw new Error('ไม่พบคำขอ: '+requestId);
  if(String(row[statusCol]||'')!=='PENDING') throw new Error('คำขอนี้ถูกดำเนินการแล้ว');
  setCell_(row,rh,'Status','REJECTED'); setCell_(row,rh,'Reviewed At',new Date()); setCell_(row,rh,'Reviewed By','Admin'); setCell_(row,rh,'Admin Reason',reason);
  rs.getRange(rowNo,1,1,rh.length).setValues([row]); SpreadsheetApp.flush(); return {ok:true,requestId};
}
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
      let newStock=toNumber_(found.quantity), rollback=''; if(type==='เบิก'){newStock+=qty;rollback='ปรับเพิ่ม';} else if(type==='รับเข้า'){if(newStock<qty)throw new Error('จำนวนในคลังไม่พอสำหรับย้อนผลผลิต: '+found.name);newStock-=qty;rollback='ปรับลด';} else if(type==='คืน'){if(newStock<qty)throw new Error('จำนวนในคลังไม่พอสำหรับย้อนการคืน: '+found.name);newStock-=qty;rollback='ปรับลด';} else continue;
      inv.getRange(found.row,found.quantityCol).setValue(newStock);
      const txId=nextId_(il,'Transaction ID','T',5);
      const reverseOf=type==='เบิก'?'เบิก':type==='คืน'?'คืน':type==='รับเข้า'?'รับเข้า':type;
      appendInventoryLog_(il,lh,{txId,dateValue:new Date(),time:Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'HH:mm'),type:rollback,itemId:found.id,name:found.name,qty:qty,unit:found.unit,workId:workId,recorder:email,note:'ยกเลิก '+workId+' | ย้อนรายการ: '+reverseOf+' | '+reason});
      reverted.push({item:found.name,quantity:qty,unit:found.unit,type:rollback});
    }
    // เก็บสำเนาไว้ใน DELETED WORK LOG ก่อนลบจริง เพื่อให้ ADMIN ยังดูประวัติย้อนหลังได้
    const deletedSheet=getOrCreateSheet_(ss,'DELETED WORK LOG');
    const deletedHeaders=ensureHeaders_(deletedSheet, headers.concat(['Deleted At','Deleted By','Delete Reason','Delete Source','Delete Request ID']));
    const deletedRow=blankRow_(deletedHeaders.length);
    headers.forEach((h,i)=>{ if(i<row.length) deletedRow[i]=row[i]; });
    setCell_(deletedRow, deletedHeaders, 'Deleted At', new Date());
    setCell_(deletedRow, deletedHeaders, 'Deleted By', email);
    setCell_(deletedRow, deletedHeaders, 'Delete Reason', reason);
    setCell_(deletedRow, deletedHeaders, 'Delete Source', data.deleteSource || 'ADMIN DIRECT');
    setCell_(deletedRow, deletedHeaders, 'Delete Request ID', data.requestId || '');
    deletedSheet.appendRow(deletedRow);

    // ลบรายการ WORK LOG ออกจากฐานข้อมูลจริง หลังย้อนรายการคลังและเก็บประวัติแล้ว
    ws.deleteRow(rowNo); SpreadsheetApp.flush();
    return {ok:true,workId:workId,deleted:true,cancelledBy:email,reverted:reverted};
  } finally { lock.releaseLock(); }
}
