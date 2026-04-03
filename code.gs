
function doGet(e) {
  const action = e.parameter.action;
  let result;

  try {
    if (action === 'getConfig') {
      result = getConfig();
    } else if (action === 'getKnownFaces') {
      result = getKnownFaces();
    } else if (action === 'getAttendance') {
      result = getAttendance();
    } else {
      result = { error: 'Unknown action: ' + action };
    }
  } catch (err) {
    logEvent('ERROR', action, '', err.message || JSON.stringify(err));
    result = { error: 'Server error: ' + (err.message || 'unknown') };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: 'Invalid JSON body' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const action = data.action;
  Logger.log('doPost received action: ' + action + ', data keys: ' + Object.keys(data).join(','));

  let result;

  try {
    if (action === 'registerUser') {
      result = registerUser(data.name, data.faceDescriptor);
    } else if (action === 'logAttendance') {
      result = logAttendance(data.name, data.lat, data.lng, data.time, data.reason);
    } else if (action === 'saveConfig') {
      result = saveConfig(data.lat, data.lng, data.radius, data.startTime, data.lateThreshold);
    } else {
      result = { error: 'Unknown action: ' + action };
    }
  } catch (err) {
    logEvent('ERROR', action, data.name || '', err.message || JSON.stringify(err));
    result = { error: 'Server error: ' + (err.message || 'unknown') };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function logEvent(level, action, name, message, payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Logs');
  if (!sheet) {
    sheet = ss.insertSheet('Logs');
    sheet.appendRow(['Timestamp', 'Level', 'Action', 'Name', 'Message', 'Payload']);
    sheet.getRange('A1:F1').setFontWeight('bold').setBackground('#f0f0f0');
    sheet.setColumnWidth(1, 150);
    sheet.setColumnWidth(2, 80);
    sheet.setColumnWidth(3, 100);
    sheet.setColumnWidth(4, 150);
    sheet.setColumnWidth(5, 200);
    sheet.setColumnWidth(6, 200);
  }

  const now = new Date();
  sheet.appendRow([now, level, action, name, message, JSON.stringify(payload || {})]);
}

function haversineKm(lat1, lon1, lat2, lon2) {
  const R = 6371;
  const dLat = (lat2 - lat1) * Math.PI / 180;
  const dLon = (lon2 - lon1) * Math.PI / 180;
  const a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
            Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) *
            Math.sin(dLon / 2) * Math.sin(dLon / 2);
  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  return R * c;
}

function getAttendance() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Attendance');
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const header = data[0];
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const row = {};
    header.forEach((col, j) => {
      row[col] = data[i][j];
    });
    rows.push(row);
  }
  return rows;
}

// --- ส่วนจัดการใบหน้า (Users) ---
function registerUser(name, faceDescriptor) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Users');
  if (!sheet) {
    sheet = ss.insertSheet('Users');
    sheet.appendRow(['Name', 'Face Descriptor', 'Registered At']);
    sheet.getRange('A1:C1').setFontWeight('bold').setBackground('#f0f0f0');
    sheet.setColumnWidth(1, 150);
    sheet.setColumnWidth(2, 300);
    sheet.setColumnWidth(3, 150);
  }

  if (!name || !faceDescriptor || !Array.isArray(faceDescriptor) || faceDescriptor.length !== 128) {
    logEvent('WARN', 'registerUser', name || '', 'Invalid user data', {faceDescriptorLength: faceDescriptor ? faceDescriptor.length : 'n/a'});
    return { error: 'ข้อมูลไม่ถูกต้อง: ชื่อหรือหน้าไม่สมบูรณ์' };
  }

  // ตรวจซ้ำชื่อ
  const existing = sheet.getDataRange().getValues().slice(1).find(r => r[0] === name);
  if (existing) {
    return { success: false, message: 'ชื่อผู้ใช้นี้มีอยู่แล้ว' };
  }

  sheet.appendRow([name, JSON.stringify(faceDescriptor), new Date()]);
  Logger.log('User registered: ' + name);
  logEvent('INFO', 'registerUser', name, 'Registered face user');
  return { success: true, message: 'บันทึกข้อมูลหน้าเรียบร้อย' };
}

function getKnownFaces() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Users');
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  let users = [];
  for (let i = 1; i < data.length; i++) {
    const name = data[i][0];
    const jsonStr = data[i][1];
    if (name && jsonStr) {
      try {
        users.push({ label: name, descriptor: JSON.parse(jsonStr) });
      } catch (e) {}
    }
  }
  return users;
}

// --- ส่วนบันทึกเวลา (Attendance) ---
function logAttendance(name, lat, lng, time, reason) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Attendance');
  if (!sheet) {
    sheet = ss.insertSheet('Attendance');
    sheet.appendRow(['Name', 'Time', 'Date', 'Latitude', 'Longitude', 'Google Map Link', 'Status', 'DistanceKM', 'Reason']);
    sheet.getRange('A1:I1').setFontWeight('bold').setBackground('#f0f0f0');
    sheet.setColumnWidth(1, 150);
    sheet.setColumnWidth(2, 100);
    sheet.setColumnWidth(3, 100);
    sheet.setColumnWidth(4, 100);
    sheet.setColumnWidth(5, 100);
    sheet.setColumnWidth(6, 200);
    sheet.setColumnWidth(7, 100);
    sheet.setColumnWidth(8, 100);
    sheet.setColumnWidth(9, 150);
  const conditionalFormatRules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('late')
      .setBackground('#ffcccc')
      .setRanges([sheet.getRange('H2:H')])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('on-time')
      .setBackground('#ccffcc')
      .setRanges([sheet.getRange('H2:H')])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('early')
      .setBackground('#ffffcc')
      .setRanges([sheet.getRange('H2:H')])
      .build()
  ];
  sheet.setConditionalFormatRules(conditionalFormatRules);
  }

  const config = getConfig();
  const now = time ? new Date(time) : new Date();
  const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'd/M/yyyy');
  const timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss');
  const mapLink = (lat && lng) ? `https://www.google.com/maps?q=${lat},${lng}` : '';

  let status = 'unknown';
  let distance = 0;

  if (lat && lng && config.lat && config.lng) {
    distance = haversineKm(parseFloat(lat), parseFloat(lng), config.lat, config.lng);
    if (isNaN(distance)) distance = 0;

    if (distance > config.radius) {
      status = 'out-of-zone';
    } else {
      const timeValue = timeStr;
      const target = config.startTime || '09:00:00';
      const lateBound = config.lateThreshold || 10;
      const minute = function (t) { const p=t.split(':'); return Number(p[0])*60 + Number(p[1]); };
      const nowMin = minute(timeValue);
      const targetMin = minute(target);

      if (nowMin < targetMin) {
        status = 'early';
      } else if (nowMin <= targetMin + lateBound) {
        status = 'on-time';
      } else {
        status = 'late';
      }
    }
  } else {
    status = 'missing-location';
  }

  sheet.appendRow([name, timeStr, "'" + dateStr, lat || '-', lng || '-', mapLink, status, distance.toFixed(3), reason || '']);
  Logger.log('Attendance recorded for ' + name + ' with status ' + status + ' reason: ' + reason);
  logEvent('INFO', 'logAttendance', name, 'Attendance recorded', {status: status, distanceKM: distance, reason: reason});

  return { success: true, message: 'บันทึกเวลาสำเร็จ', status: status, distance: distance };
}

// --- ส่วนจัดการ Config (GPS) ---
function saveConfig(lat, lng, radius, startTime, lateThreshold) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Config');

  if (!sheet) {
    sheet = ss.insertSheet('Config');
    sheet.getRange('A1:B1').setValues([['Parameter', 'Value']]);
    sheet.getRange('A2').setValue('Target Latitude');
    sheet.getRange('A3').setValue('Target Longitude');
    sheet.getRange('A4').setValue('Allowed Radius (KM)');
    sheet.getRange('A5').setValue('Work Start Time (HH:MM:SS)');
    sheet.getRange('A6').setValue('Late Threshold (minute)');
    sheet.setColumnWidth(1, 200);
    sheet.getRange('A1:B1').setFontWeight('bold').setBackground('#f0f0f0');
    sheet.getRange('A2:A6').setFontWeight('bold');
  }

  sheet.getRange('B2').setValue(lat);
  sheet.getRange('B3').setValue(lng);
  sheet.getRange('B4').setValue(radius);
  sheet.getRange('B5').setValue(startTime || '08:00:00');
  sheet.getRange('B6').setValue(lateThreshold || 10);

  Logger.log('Config saved');
  logEvent('INFO', 'saveConfig', '', 'Config saved', {lat, lng, radius, startTime, lateThreshold});
  return { success: true, message: 'บันทึกการตั้งค่าลง Google Sheets เรียบร้อย' };
}

function getConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Config');

  let config = { lat: 0, lng: 0, radius: 0.5, startTime: '08:00:00', lateThreshold: 10 };

  if (sheet) {
    const latVal = sheet.getRange('B2').getValue();
    const lngVal = sheet.getRange('B3').getValue();
    const radiusVal = sheet.getRange('B4').getValue();
    const startTimeVal = sheet.getRange('B5').getValue();
    const lateThresholdVal = sheet.getRange('B6').getValue();

    if (latVal !== '') config.lat = parseFloat(latVal);
    if (lngVal !== '') config.lng = parseFloat(lngVal);
    if (radiusVal !== '') config.radius = parseFloat(radiusVal);
    if (startTimeVal) config.startTime = startTimeVal;
    if (lateThresholdVal !== '') config.lateThreshold = parseInt(lateThresholdVal, 10);
  }

  return config;
}