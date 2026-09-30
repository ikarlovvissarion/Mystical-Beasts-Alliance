const SPREADSHEET_ID = '1eph0zMO0iEtgqSsGEVLVOr9txgx9J2bAZyPOuwvmBh8';

// แก้รายการนี้ได้ตามปฏิทินของชมรม หาก "วันผู้วิเศษ" ใช้ชื่อเฉพาะของเซิร์ฟเวอร์
const MAGICAL_DAYS = [
  'วันจันทร์', 'วันอังคาร', 'วันพุธ', 'วันพฤหัสบดี',
  'วันศุกร์', 'วันเสาร์', 'วันอาทิตย์'
];

const WORK_TYPES = {
  'งานทั่วไป': [
    'การทำความสะอาด',
    'การให้อาหารสัตว์วิเศษ',
    'การรดน้ำพืชผัก',
    'การตรวจสุขภาพประจำวันสัตว์วิเศษ',
    'การรักษาสัตว์วิเศษในคลินิกสัตว์วิเศษ'
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
      'เบิกคลังชมรม', 'จำนวนที่เบิก', 'หน่วยที่เบิก', 'จำนวนที่ได้รับ', 'หน่วยที่ได้รับ'
    ]);

    const inventoryHeaders = ensureHeaders_(inventorySheet, ['Category', 'ID', 'รายการ', 'จำนวน', 'หน่วย']);
    const inventoryLogHeaders = ensureHeaders_(inventoryLogSheet, [
      'Transaction ID', 'วันที่ (มักเกิ้ล)', 'เวลา (IC)', 'ประเภท',
      'Item ID', 'รายการ', 'จำนวน', 'หน่วย', 'Work ID', 'ผู้ลงบันทึก', 'หมายเหตุ'
    ]);

    const workId = nextId_(workSheet, 'Work ID', 'W', 4);
    const now = new Date();
    const dateValue = data.muggleDate ? new Date(data.muggleDate + 'T00:00:00') : now;

    // ตรวจสอบรายการคลังก่อนเขียน WORK LOG เพื่อไม่ให้เกิด log ค้างเมื่อยอดไม่พอ
    let withdrawItem = null;
    let receiveItem = null;
    if (data.category === 'งานทั่วไป' && data.withdrawItemName && Number(data.withdrawQty) > 0) {
      withdrawItem = findInventoryItem_(inventorySheet, inventoryHeaders, data.withdrawItemId, data.withdrawItemName);
      if (!withdrawItem) throw new Error('ไม่พบรายการที่ต้องการเบิกใน INVENTORY');
      if (toNumber_(withdrawItem.quantity) < Number(data.withdrawQty)) throw new Error('จำนวนในคลังไม่เพียงพอ');
    }
    if (data.category === 'เก็บผลผลิต' && Number(data.receivedQty) > 0) {
      receiveItem = findInventoryItem_(inventorySheet, inventoryHeaders, data.receivedItemId, data.receivedItemName);
      if (!receiveItem) throw new Error('ไม่พบรายการผลผลิตใน INVENTORY');
    }

    const row = blankRow_(workHeaders.length);
    setCell_(row, workHeaders, 'Category', data.category);
    setCell_(row, workHeaders, 'Work ID', workId);
    setCell_(row, workHeaders, 'วันที่ (มักเกิ้ล)', dateValue);
    setCell_(row, workHeaders, 'วันผู้วิเศษ', data.magicalDay);
    setCell_(row, workHeaders, 'เวลา (IC)', data.icTime);
    setCell_(row, workHeaders, 'ประเภทการทำงาน', data.workType);
    setCell_(row, workHeaders, 'Target ID', data.targetId);
    setCell_(row, workHeaders, 'Target Name', data.targetName);
    setCell_(row, workHeaders, 'ผู้ลงบันทึก', data.recorder);
    setCell_(row, workHeaders, 'เบิกคลังชมรม', data.withdrawItemName || '');
    setCell_(row, workHeaders, 'จำนวนที่เบิก', data.withdrawQty || '');
    setCell_(row, workHeaders, 'หน่วยที่เบิก', data.withdrawUnit || '');
    setCell_(row, workHeaders, 'จำนวนที่ได้รับ', data.receivedQty || '');
    setCell_(row, workHeaders, 'หน่วยที่ได้รับ', data.receivedUnit || '');
    workSheet.appendRow(row);

    const transactions = [];

    if (withdrawItem) {
      const qty = Number(data.withdrawQty);
      const stock = toNumber_(withdrawItem.quantity);
      const unit = withdrawItem.unit || data.withdrawUnit || '';
      inventorySheet.getRange(withdrawItem.row, withdrawItem.quantityCol).setValue(stock - qty);
      const txId = nextId_(inventoryLogSheet, 'Transaction ID', 'T', 5);
      appendInventoryLog_(inventoryLogSheet, inventoryLogHeaders, {
        txId, dateValue, time: data.icTime, type: 'เบิก', itemId: withdrawItem.id,
        name: withdrawItem.name, qty, unit, workId, recorder: data.recorder,
        note: data.workType
      });
      transactions.push({ id: txId, type: 'เบิก', item: withdrawItem.name, quantity: qty, unit });
    }

    if (receiveItem) {
      const qty = Number(data.receivedQty);
      const stock = toNumber_(receiveItem.quantity);
      const unit = receiveItem.unit || data.receivedUnit || '';
      inventorySheet.getRange(receiveItem.row, receiveItem.quantityCol).setValue(stock + qty);
      const txId = nextId_(inventoryLogSheet, 'Transaction ID', 'T', 5);
      appendInventoryLog_(inventoryLogSheet, inventoryLogHeaders, {
        txId, dateValue, time: data.icTime, type: 'รับเข้า', itemId: receiveItem.id,
        name: receiveItem.name, qty, unit, workId, recorder: data.recorder,
        note: data.workType
      });
      transactions.push({ id: txId, type: 'รับเข้า', item: receiveItem.name, quantity: qty, unit });
    }

    SpreadsheetApp.flush();
    return { ok: true, workId, transactions };
  } finally {
    lock.releaseLock();
  }
}

function validate_(data) {
  if (!data.category || !['งานทั่วไป', 'เก็บผลผลิต'].includes(data.category)) throw new Error('Category ไม่ถูกต้อง');
  if (!data.workType || !WORK_TYPES[data.category].includes(data.workType)) throw new Error('ประเภทการทำงานไม่ถูกต้อง');
  if (!data.targetId) throw new Error('กรุณาเลือก Target ID');
  if (!data.targetName) throw new Error('ไม่พบ Target Name');
  if (!data.recorder) throw new Error('กรุณาเลือกผู้ลงบันทึก');
  if (!data.muggleDate) throw new Error('กรุณาเลือกวันที่');
  if (!data.magicalDay) throw new Error('กรุณาเลือกวันผู้วิเศษ');
  if (!data.icTime) throw new Error('กรุณาเลือกเวลา IC');
}

function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function ensureHeaders_(sheet, expected) {
  const lastCol = Math.max(sheet.getLastColumn(), expected.length);
  const existing = sheet.getLastRow() ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String) : [];
  if (!sheet.getLastRow()) {
    sheet.getRange(1, 1, 1, expected.length).setValues([expected]);
    return expected;
  }
  const headers = existing.slice(0, Math.max(existing.length, expected.length));
  expected.forEach((h, i) => {
    if (!headers[i]) headers[i] = h;
  });
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  return headers;
}

function blankRow_(length) { return Array.from({ length }, () => ''); }
function setCell_(row, headers, name, value) {
  const i = headers.indexOf(name);
  if (i >= 0) row[i] = value;
}
function findHeader_(headers, names) {
  for (const n of names) {
    const i = headers.indexOf(n);
    if (i >= 0) return i;
  }
  return -1;
}

function findInventoryItem_(sheet, headers, id, name) {
  const idCol = findHeader_(headers, ['ID', 'Item ID']);
  const nameCol = findHeader_(headers, ['รายการ', 'Name', 'Item']);
  const qtyCol = findHeader_(headers, ['จำนวน', 'Quantity']);
  const unitCol = findHeader_(headers, ['หน่วย', 'Unit']);
  if (nameCol < 0 || qtyCol < 0) return null;
  const values = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 0), headers.length).getValues();
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
  setCell_(row, headers, 'Transaction ID', tx.txId);
  setCell_(row, headers, 'วันที่ (มักเกิ้ล)', tx.dateValue);
  setCell_(row, headers, 'เวลา (IC)', tx.time);
  setCell_(row, headers, 'ประเภท', tx.type);
  setCell_(row, headers, 'Item ID', tx.itemId);
  setCell_(row, headers, 'รายการ', tx.name);
  setCell_(row, headers, 'จำนวน', tx.qty);
  setCell_(row, headers, 'หน่วย', tx.unit);
  setCell_(row, headers, 'Work ID', tx.workId);
  setCell_(row, headers, 'ผู้ลงบันทึก', tx.recorder);
  setCell_(row, headers, 'หมายเหตุ', tx.note || '');
  sheet.appendRow(row);
}

function nextId_(sheet, headerName, prefix, width) {
  const headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0].map(String);
  const col = headers.indexOf(headerName);
  if (col < 0 || sheet.getLastRow() < 2) return prefix + String(1).padStart(width, '0');
  const values = sheet.getRange(2, col + 1, sheet.getLastRow() - 1, 1).getValues().flat();
  let max = 0;
  values.forEach(v => {
    const m = String(v || '').match(/(\d+)$/);
    if (m) max = Math.max(max, Number(m[1]));
  });
  return prefix + String(max + 1).padStart(width, '0');
}

function toNumber_(v) {
  if (typeof v === 'number') return v;
  const n = Number(String(v || '').replace(/,/g, '').trim());
  return Number.isFinite(n) ? n : 0;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
