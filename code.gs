const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

function doGet(e) {
  const action = e.parameter.action;

  if (action === 'search') {
    return jsonOutput(searchRows(e.parameter.sheet, e.parameter.keyword || ''));
  }

  if (action === 'list') {
    return jsonOutput(listRows(e.parameter.sheet));
  }

  if (action === 'today') {
    return jsonOutput(getTodayFollowUps());
  }

  return jsonOutput({ ok: false, error: '未知的讀取動作' });
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents || '{}');
  const action = body.action;

  if (action === 'add') {
    return jsonOutput(addRow(body.sheet, body.data));
  }

  if (action === 'checkPhone') {
    return jsonOutput(checkPhone(body.phone));
  }

  if (action === 'updateStatus') {
    return jsonOutput(updateStatus(body));
  }

  return jsonOutput({ ok: false, error: '未知的寫入動作' });
}

function jsonOutput(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet(sheetName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('找不到分頁：' + sheetName);
  return sheet;
}

function getHeaders(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

function addRow(sheetName, data) {
  const sheet = getSheet(sheetName);
  const headers = getHeaders(sheet);

  if (headers.includes('建立日期') && !data['建立日期']) {
    data['建立日期'] = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd');
  }

  const row = headers.map(header => data[header] || '');
  sheet.appendRow(row);

  const rowNumber = sheet.getLastRow();
  sheet.getRange(rowNumber, 1, 1, headers.length).setHorizontalAlignment('center');

if (sheetName === '客戶資料表') {
  const noteColumn = headers.indexOf('備註') + 1;
  const debugInfo = [];

  try {
    const c1 = createFollowUpCalendarEvents(data);
    const c2 = createAppointmentCalendarEvent(data);
    const c3 = createOtherReminderCalendarEvent(data);
    debugInfo.push('收到提醒值=' + (data['Google日曆提醒時間'] || '(空)'));
    debugInfo.push('追蹤建立=' + c1 + '筆');
    debugInfo.push('預約建立=' + c2 + '筆');
    debugInfo.push('其它建立=' + c3 + '筆');

    if (noteColumn > 0) {
      sheet.getRange(rowNumber, noteColumn).setValue(debugInfo.join(' / '));
    }
  } catch (err) {
    if (noteColumn > 0) {
      sheet.getRange(rowNumber, noteColumn).setValue('日曆錯誤：' + err.message + ' / ' + debugInfo.join(' / '));
    }
    throw err;
  }
}
  return { ok: true };
}

function listRows(sheetName) {
  const sheet = getSheet(sheetName);
  const headers = getHeaders(sheet);
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return { ok: true, data: [] };

  const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();

  const data = values.map((row, index) => {
    const item = {};
    headers.forEach((header, i) => item[header] = row[i]);
    item._rowNumber = index + 2;
    return item;
  });

  return { ok: true, data };
}

function createFollowUpCalendarEvents(data){
  const startText = data['Google日曆提醒時間'];
  if (!startText) return 0;

  const customerName = data['客戶姓名'] || '客戶';
  const trackType = data['追蹤區分'] && data['追蹤區分'] !== '未選擇'
    ? data['追蹤區分']
    : '追蹤';

  let count = 0;

  [
  { label: 1, offset: 1 },
  { label: 4, offset: 4 },
  { label: 6, offset: 6 }
].forEach(item => {
  const start = parseDateTime(startText);
  if (isNaN(start.getTime())) return;

  start.setDate(start.getDate() + item.offset);
    const eventStart = findAvailableCalendarTime(start, 30);
    const eventEnd = new Date(eventStart.getTime() + 30 * 60 * 1000);

    const event = CalendarApp.getDefaultCalendar().createEvent(
      `${customerName}｜${trackType}第${item.label}日提醒`,
      eventStart,
      eventEnd,
      { description: buildCalendarDescription(data) }
    );

    event.setColor(CalendarApp.EventColor.GREEN);
    event.addPopupReminder(30);
    event.addPopupReminder(60 * 24);
    count++;
  });

  return count;
}
function createOtherReminderCalendarEvent(data) {
  const startText = data['其它提醒日期時間'];
  if (!startText) return 0;

  const start = parseDateTime(startText);
  if (isNaN(start.getTime())) return 0;

  const customerName = data['客戶姓名'] || '客戶';

  const eventStart = findAvailableCalendarTime(start, 30);
  const eventEnd = new Date(eventStart.getTime() + 30 * 60 * 1000);

  const event = CalendarApp.getDefaultCalendar().createEvent(
    `${customerName}｜其它提醒`,
    eventStart,
    eventEnd,
    { description: buildCalendarDescription(data) }
  );

  event.setColor(CalendarApp.EventColor.GREEN);
  event.addPopupReminder(30);
  event.addPopupReminder(60 * 24);
  return 1;
}

function createAppointmentCalendarEvent(data) {
  const startText = data['預約日期時間'];
  if (!startText) return 0;

 const start = parseDateTime(startText);
  if (isNaN(start.getTime())) return 0;

  const customerName = data['客戶姓名'] || '客戶';
  const appointmentType = data['預約類型'] && data['預約類型'] !== '未選擇'
    ? data['預約類型']
    : '預約';

  const eventStart = findAvailableCalendarTime(start, 60);
  const eventEnd = new Date(eventStart.getTime() + 60 * 60 * 1000);

  const event = CalendarApp.getDefaultCalendar().createEvent(
    `${customerName}｜${appointmentType}`,
    eventStart,
    eventEnd,
    { description: buildCalendarDescription(data) }
  );

  event.setColor(CalendarApp.EventColor.GREEN);
  event.addPopupReminder(30);
  event.addPopupReminder(60 * 24);
  return 1;
}

function findAvailableCalendarTime(start, minutes) {
  const calendar = CalendarApp.getDefaultCalendar();
  let current = new Date(start);

  for (let i = 0; i < 12; i++) {
    const end = new Date(current.getTime() + minutes * 60 * 1000);
    const events = calendar.getEvents(current, end);

    if (events.length === 0) return current;

    current = new Date(current.getTime() + minutes * 60 * 1000);
  }

  return current;
}

function buildCalendarDescription(data) {
  return [
    data['電話'] ? `電話：${data['電話']}` : '',
    data['LINE名稱'] ? `LINE：${data['LINE名稱']}` : '',
    data['購買方案'] ? `購買方案：${data['購買方案']}` : '',
    data['後援人'] ? `後援人：${data['後援人']}` : '',
    data['排線'] ? `排線：${data['排線']}` : '',
    data['備註'] ? `備註：${data['備註']}` : ''
  ].filter(Boolean).join('\n');
}

function searchRows(sheetName, keyword) {
  const result = listRows(sheetName);
  if (!result.ok) return result;

  const kw = String(keyword || '').trim().toLowerCase();

  if (!kw) return result;

  const data = result.data.filter(row => {
    return Object.values(row).some(value =>
      String(value || '').toLowerCase().includes(kw)
    );
  });

  return { ok: true, data };
}

function checkPhone(phone) {
  const result = listRows('客戶資料表');
  if (!result.ok) return result;

  const target = String(phone || '').trim();

  const data = result.data.filter(row =>
    String(row['電話'] || '').trim() === target
  );

  return { ok: true, data };
}

function getTodayFollowUps() {
  const result = listRows('追蹤紀錄表');
  if (!result.ok) return result;

  const today = new Date();
  today.setHours(23, 59, 59, 999);

  const data = result.data.filter(row => {
    const done = String(row['是否完成'] || '').trim();
    const followDate = row['追蹤日期時間'];

    if (done === '是') return false;
    if (!followDate) return true;

    const date = new Date(followDate);
    if (isNaN(date.getTime())) return true;

    return date <= today;
  }).map(row => {
    row['追蹤ID'] = row._rowNumber;
    return row;
  });

  return { ok: true, data };
}

function updateStatus(body) {
  const sheet = getSheet(body.sheet);
  const headers = getHeaders(sheet);
  const updates = body.updates || {};

  let rowNumber = Number(body.id);

  if (!rowNumber && body.idField) {
    const idColumnIndex = headers.indexOf(body.idField);
    if (idColumnIndex >= 0) {
      const values = sheet.getRange(2, idColumnIndex + 1, sheet.getLastRow() - 1, 1).getValues();
      const foundIndex = values.findIndex(row => String(row[0]) === String(body.id));
      if (foundIndex >= 0) rowNumber = foundIndex + 2;
    }
  }

  if (!rowNumber || rowNumber < 2) {
    return { ok: false, error: '找不到要更新的資料列' };
  }

  Object.keys(updates).forEach(key => {
    const columnIndex = headers.indexOf(key);
    if (columnIndex >= 0) {
      sheet.getRange(rowNumber, columnIndex + 1).setValue(updates[key]);
    }
  });

  return { ok: true };
}

function parseDateTime(value) {
  if (!value) return new Date('');

  if (value instanceof Date) return value;

  const text = String(value).trim().replace(/\//g, '-').replace('T', ' ');
  const parts = text.split(' ');
  const datePart = parts[0];
  const timePart = parts[1] || '10:00';

  const datePieces = datePart.split('-').map(Number);
  const timePieces = timePart.split(':').map(Number);

  return new Date(
    datePieces[0],
    datePieces[1] - 1,
    datePieces[2],
    timePieces[0] || 10,
    timePieces[1] || 0
  );
}
