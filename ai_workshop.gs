// ======================================================
// ตั้งค่า ID
// ======================================================
var FOLDER_ID  = '18ROY7lBopPlJ25HdeUhkNxPCy7eYiVgj';
var SHEET_ID   = '1TPtqTluS4zrRtCKaa_M3WcbZYyJAsGZCJ2u7lVFZWzk';
var SHEET_NAME = 'Sheet1';

// อีเมลรับแจ้งเตือนเมื่อมีผู้สมัครใหม่
var NOTIFY_EMAIL = 'tontoomgracesean@gmail.com';

// LINE OA Messaging API
var LINE_TOKEN   = 'DteOwBMd5u1u46a7ehVZJ0OwU1HJ9UFq3KsMbh5ZDtAfSjtFTyRw8PLqnMkv/PIDeqZtDHgHFdv0gYlYibxd0kI5QCpFS516YUnluhTDj0h27zQjzNtcbnYU8CZrohdWXCFxpzVz2z9IYRYicq0NugdB04t89/1O/w1cDnyilFU=';
var LINE_USER_ID = 'Ua4ddc5785f2f935897ed131c1058a8d1';

// ======================================================
// Router — รับ GET request
// ======================================================
function doGet(e) {
  var action = e.parameter.action;
  if (action == 'getCounts')      return getBookingCountsJSON();
  if (action == 'login')          return handleLogin(e);
  if (action == 'dashboard')      return handleDashboard(e);
  if (action == 'updateStatus')   return handleUpdateStatus(e);
  return ContentService.createTextOutput('System is running correctly (API Mode).');
}

// ======================================================
// นับจำนวนผู้สมัครแต่ละหลักสูตร
// ======================================================
function getBookingCountsJSON() {
  try {
    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    if (!sheet) return responseJSON({ medical: 0, research: 0, error: 'Sheet Not Found' });

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return responseJSON({ medical: 0, research: 0 });

    var data   = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    var counts = { medical: 0, research: 0 };
    for (var i = 0; i < data.length; i++) {
      var c = String(data[i][0]);
      if (c.indexOf('AI การแพทย์') !== -1)  counts.medical++;
      else if (c.indexOf('AI งานวิจัย') !== -1) counts.research++;
    }
    return responseJSON(counts);
  } catch (err) {
    return responseJSON({ medical: 0, research: 0, error: err.toString() });
  }
}

// ======================================================
// Helper — สร้าง JSON Response
// ======================================================
function responseJSON(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ======================================================
// Login — ตรวจสอบ credentials แล้วคืน token
// (ไฟล์นี้เป็น server-side เท่านั้น browser อ่านไม่ได้)
// ======================================================
function handleLogin(e) {
  var u = e.parameter.u || '';
  var p = e.parameter.p || '';

  // ตรวจสอบ credentials
  var validUser = 'tontoomgracesean@gmail.com';
  var validPass = 'Ton@0821709518';

  if (u !== validUser || p !== validPass) {
    return responseJSON({ ok: false });
  }

  // สร้าง token อายุ 8 ชั่วโมง
  var token = generateToken_();
  var expiry = new Date().getTime() + (8 * 60 * 60 * 1000);
  var props  = PropertiesService.getScriptProperties();
  props.setProperty('DASH_TOKEN',  token);
  props.setProperty('DASH_EXPIRY', String(expiry));

  return responseJSON({ ok: true, token: token });
}

function generateToken_() {
  var raw = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    new Date().toString() + String(Math.random())
  );
  return raw.map(function(b) {
    return (b < 0 ? b + 256 : b).toString(16).padStart(2, '0');
  }).join('').substring(0, 40);
}

function validateToken_(token) {
  if (!token) return false;
  var props  = PropertiesService.getScriptProperties();
  var stored = props.getProperty('DASH_TOKEN')  || '';
  var expiry = parseInt(props.getProperty('DASH_EXPIRY') || '0');
  return token === stored && new Date().getTime() < expiry;
}

// ======================================================
// Dashboard — ดึงข้อมูลผู้สมัครทั้งหมด
// ======================================================
function handleDashboard(e) {
  var token = e.parameter.token || '';
  if (!validateToken_(token)) return responseJSON({ ok: false, error: 'Unauthorized' });

  try {
    var sheet   = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return responseJSON({ ok: true, rows: [], medical: 0, research: 0 });

    var data   = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
    var rows   = [];
    var counts = { medical: 0, research: 0 };

    for (var i = 0; i < data.length; i++) {
      var course = String(data[i][1]);
      if (course.indexOf('AI การแพทย์') !== -1)      counts.medical++;
      else if (course.indexOf('AI งานวิจัย') !== -1) counts.research++;

      var rawDate = data[i][0];
      var dateStr = rawDate ? new Date(rawDate).toLocaleDateString('th-TH') : '-';

      rows.push({
        rowNum:  i + 2,
        date:    dateStr,
        course:  course,
        name:    String(data[i][2] || ''),
        phone:   String(data[i][3] || ''),
        email:   String(data[i][4] || ''),
        line:    String(data[i][5] || ''),
        org:     String(data[i][6] || ''),
        expect:  String(data[i][7] || ''),
        status:  String(data[i][9] || 'รอตรวจสอบ')
      });
    }

    return responseJSON({ ok: true, rows: rows, medical: counts.medical, research: counts.research });
  } catch (err) {
    return responseJSON({ ok: false, error: err.toString() });
  }
}

// ======================================================
// อัปเดตสถานะในชีต (เรียกจาก Dashboard)
// ======================================================
function handleUpdateStatus(e) {
  var token  = e.parameter.token  || '';
  if (!validateToken_(token)) return responseJSON({ ok: false, error: 'Unauthorized' });

  var rowNum    = parseInt(e.parameter.row    || '0');
  var newStatus = e.parameter.status || 'อนุมัติแล้ว';

  if (rowNum < 2) return responseJSON({ ok: false, error: 'Invalid row' });

  try {
    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    sheet.getRange(rowNum, 10).setValue(newStatus);

    // แจ้งเตือน Line OA เมื่ออนุมัติ
    if (newStatus === 'อนุมัติแล้ว') {
      var row = sheet.getRange(rowNum, 1, 1, 10).getValues()[0];
      sendApprovalNotification_(row);
    }

    return responseJSON({ ok: true });
  } catch (err) {
    return responseJSON({ ok: false, error: err.toString() });
  }
}

// ======================================================
// แจ้งเตือนเมื่ออนุมัติผู้สมัคร — ส่งทั้ง LINE OA + Email
// ======================================================
function sendApprovalNotification_(row) {
  var course = String(row[1] || '-');
  var name   = String(row[2] || '-');
  var phone  = String(row[3] || '-');
  var email  = String(row[4] || '-');
  var lineId = String(row[5] || '-');
  var org    = String(row[6] || '-');
  var expect = String(row[7] || '-').substring(0, 120);

  var lineMsg = '👤 ชื่อ: '    + name   + '\n' +
                '📞 เบอร์: '   + phone  + '\n' +
                '💬 Line ID: ' + lineId + '\n' +
                '🏢 หน่วยงาน: '+ org    + '\n' +
                'อนุมัติแล้ว ✅';

  // LINE OA broadcast → ส่งหาทุกคนที่ follow OA (ครั้งเดียว)
  try {
    var res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/broadcast', {
      method:             'POST',
      contentType:        'application/json',
      headers:            { 'Authorization': 'Bearer ' + LINE_TOKEN },
      payload:            JSON.stringify({
        messages: [{ type: 'text', text: lineMsg }]
      }),
      muteHttpExceptions: true
    });
    Logger.log('LINE approval HTTP ' + res.getResponseCode() + ': ' + res.getContentText());
  } catch (err) {
    Logger.log('LINE approval error: ' + err.toString());
  }
}

// ======================================================
// รับข้อมูลลงทะเบียน (POST) — บันทึก Sheet + อัปโหลดสลิป
// ======================================================
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    var data  = JSON.parse(e.postData.contents);

    // อัปโหลดสลิปไปยัง Google Drive
    var fileUrl = 'ไม่พบไฟล์แนบ';
    if (data.fileData && data.fileName) {
      try {
        var folder      = DriveApp.getFolderById(FOLDER_ID);
        var contentType = data.fileMimeType || 'image/jpeg';
        var blob        = Utilities.newBlob(Utilities.base64Decode(data.fileData), contentType, data.fileName);
        var file        = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        fileUrl = file.getUrl();
      } catch (err) {
        fileUrl = 'Upload Error: ' + err.toString();
      }
    }

    // บันทึกลง Sheet
    var nextRow = sheet.getLastRow() + 1;
    sheet.getRange(nextRow, 1, 1, 10).setValues([[
      new Date(),
      data['หลักสูตรที่เลือก'] || '',
      data['ชื่อ-สกุล']        || '',
      data['เบอร์โทร']          || '',
      data['Email']             || '',
      data['Line ID']           || '',
      data['หน่วยงาน']          || '',
      data['ความคาดหวัง']       || '',
      fileUrl,
      'รอตรวจสอบ'
    ]]);

    // แจ้งเตือน Line OA
    sendLineNotification_(data, fileUrl);

    return responseJSON({ result: 'success', row: nextRow });
  } catch (err) {
    return responseJSON({ result: 'error', error: err.toString() });
  } finally {
    lock.releaseLock();
  }
}

// ======================================================
// แจ้งเตือนเมื่อมีผู้สมัครใหม่ — ส่งทั้ง LINE OA + Email
// ======================================================
function sendLineNotification_(data, fileUrl) {
  var expect  = (data['ความคาดหวัง'] || '-').substring(0, 120);
  var course  = data['หลักสูตรที่เลือก'] || '-';
  var name    = data['ชื่อ-สกุล']        || '-';
  var phone   = data['เบอร์โทร']          || '-';
  var email   = data['Email']             || '-';
  var lineId  = data['Line ID']           || '-';
  var org     = data['หน่วยงาน']          || '-';

  var lineMsg = '🎉 มีผู้สมัครใหม่!\n' +
                '━━━━━━━━━━━━━━━━\n' +
                '📚 หลักสูตร: ' + course + '\n' +
                '👤 ชื่อ: '     + name   + '\n' +
                '📞 เบอร์: '    + phone  + '\n' +
                '📧 Email: '    + email  + '\n' +
                '💬 Line ID: '  + lineId + '\n' +
                '🏢 หน่วยงาน: ' + org    + '\n' +
                '💡 ความคาดหวัง: ' + expect + '\n' +
                '━━━━━━━━━━━━━━━━\n' +
                '🧾 สลิป: ' + fileUrl + '\n' +
                '━━━━━━━━━━━━━━━━\n' +
                '🌐 หน้าลงทะเบียน:\nhttps://thonganek.github.io/ai-workshop/\n' +
                '📊 Dashboard:\nhttps://thonganek.github.io/ai-workshop/ai_workshop_dashboard.html';

  // LINE OA broadcast → ส่งหาทุกคนที่ follow OA
  try {
    var res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/broadcast', {
      method:             'POST',
      contentType:        'application/json',
      headers:            { 'Authorization': 'Bearer ' + LINE_TOKEN },
      payload:            JSON.stringify({
        messages: [{ type: 'text', text: lineMsg }]
      }),
      muteHttpExceptions: true
    });
    Logger.log('LINE HTTP ' + res.getResponseCode() + ': ' + res.getContentText());
  } catch (err) {
    Logger.log('LINE error: ' + err.toString());
  }

  // Email
  try {
    var htmlBody =
      '<div style="font-family:sans-serif;max-width:560px;padding:24px;background:#f9fafb;border-radius:12px">' +
      '<h2 style="color:#1d4ed8;margin:0 0 16px">🎉 มีผู้สมัครใหม่!</h2>' +
      '<table style="width:100%;border-collapse:collapse;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,.08)">' +
      row_('หลักสูตร',    course) + row_('ชื่อ-สกุล',   name)   +
      row_('เบอร์โทร',    phone)  + row_('Email',       email)  +
      row_('Line ID',     lineId) + row_('หน่วยงาน',   org)    +
      row_('ความคาดหวัง', expect) +
      '</table>' +
      '<p style="margin:16px 0 0"><a href="' + fileUrl + '" style="background:#1d4ed8;color:#fff;padding:10px 20px;border-radius:8px;text-decoration:none;font-weight:700">ดูสลิปโอนเงิน</a></p>' +
      '</div>';
    GmailApp.sendEmail(NOTIFY_EMAIL, '🎉 มีผู้สมัครใหม่ — ' + name, lineMsg, { htmlBody: htmlBody });
  } catch (err) {
    Logger.log('Email error: ' + err.toString());
  }
}

function row_(label, value) {
  return '<tr><td style="padding:10px 14px;background:#f1f5f9;font-weight:700;width:36%;color:#374151">' +
         label + '</td><td style="padding:10px 14px;color:#111827">' + value + '</td></tr>';
}

// ======================================================
// ทดสอบทั้ง LINE OA + Email — รัน manual จาก GAS editor
// ======================================================
function testNotify() {
  var testMsg = '✅ ทดสอบแจ้งเตือน AI Workshop\nเวลา: ' + new Date();

  // LINE OA broadcast
  var lineRes = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/broadcast', {
    method:             'POST',
    contentType:        'application/json',
    headers:            { 'Authorization': 'Bearer ' + LINE_TOKEN },
    payload:            JSON.stringify({
      messages: [{ type: 'text', text: testMsg }]
    }),
    muteHttpExceptions: true
  });
  Logger.log('LINE → HTTP ' + lineRes.getResponseCode() + ': ' + lineRes.getContentText());

  // Email
  GmailApp.sendEmail(NOTIFY_EMAIL, '✅ ทดสอบแจ้งเตือน AI Workshop', testMsg);
  Logger.log('Email → ส่งไปที่ ' + NOTIFY_EMAIL + ' เรียบร้อย');
}
