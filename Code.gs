/**
 * ============================================================================
 * ระบบรายงานซ่อมบำรุงและโอเวอร์ฮอลเครื่องจักร A4 Pro (Google Apps Script Web App)
 * รองรับทั้งการเปิดหน้าเว็บแอปโดยตรง (HtmlService) และฐานข้อมูล Google Sheets & Drive
 * ============================================================================
 */

const CONFIG = {
  MAIN_SHEET_NAME: 'Maintenance_Report_Log',
  DETAIL_SHEET_NAME: 'Sections_1_to_7_Photos',
  EQUIP_SHEET_NAME: 'Equipment_Requisition_Log',
  EQUIP_ITEMS_SHEET_NAME: 'Equipment_Items_Detail',
  DRIVE_MAIN_FOLDER_NAME: 'A4_Maintenance_Report_Files',
  DEFAULT_3D_ICON_URL: 'https://cdn-icons-png.flaticon.com/512/10484/10484536.png'
};

/**
 * 1. ฟังก์ชันเริ่มต้นสำหรับกดปุ่ม "▷ เรียกใช้ (Run)" ครั้งแรกใน Apps Script
 * เพื่อสร้างตาราง Google Sheets แยกหัวข้อ 1-7, ตารางเบิก-คืนอุปกรณ์ช่าง และโฟลเดอร์ Google Drive อัตโนมัติ
 */
function setupSheetsAndDrive() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureMainSheet(ss);
  ensureDetailSheet(ss);
  ensureEquipmentSheet(ss);
  ensureEquipmentItemsSheet(ss);
  const folder = getOrCreateMainDriveFolder();

  Logger.log('✅ สร้างตารางรายงานซ่อมบำรุง & ตารางเบิก-จ่ายอุปกรณ์ช่าง สำเร็จ: ' + ss.getUrl());
  Logger.log('✅ สร้างโฟลเดอร์เก็บไฟล์รูปภาพใน Google Drive สำเร็จ: ' + folder.getUrl());
}

/**
 * 2. เปิดหน้าเว็บแอปกระดาษ A4 ทันทีเมื่อคลิกลิงก์ Web App URL (.../exec)
 * หรือดึงข้อมูล JSON เมื่อส่งพารามิเตอร์ ?action=getRecords หรือ ?action=getEquipmentRecords
 */
function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : 'webapp';

    if (action === 'getRecords') {
      const records = getAllMaintenanceRecords();
      return createJsonResponse({
        status: 'SUCCESS',
        count: records.length,
        records: records
      });
    }

    if (action === 'getEquipmentRecords') {
      const equipRecords = getAllEquipmentRecords();
      return createJsonResponse({
        status: 'SUCCESS',
        count: equipRecords.length,
        records: equipRecords
      });
    }

    const activeIconUrl = getActiveFaviconUrl();

    // ถ้ามีไฟล์ HTML ชื่อ index.html ในโปรเจ็กต์ Apps Script ให้ใช้ไฟล์นั้นเป็นหลัก
    try {
      return HtmlService.createHtmlOutputFromFile('index')
        .setTitle('MAINT & EQUIP PRO')
        .setFaviconUrl(activeIconUrl)
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=5');
    } catch (noFileErr) {
      // ถ้าวางเฉพาะไฟล์ Code.gs ไฟล์เดียว ให้แสดงหน้าเว็บแอปตัวเต็มจากฟังก์ชัน getWebAppHtml() ทันที
      return HtmlService.createHtmlOutput(getWebAppHtml())
        .setTitle('MAINT & EQUIP PRO')
        .setFaviconUrl(activeIconUrl)
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=5');
    }
  } catch (err) {
    return createJsonResponse({
      status: 'ERROR',
      message: err.toString()
    });
  }
}

/**
 * 2.1 ดึง URL ไอคอนแอปสวยๆ ที่บันทึกไว้สำหรับแสดงในหน้าต่าง "ติดตั้งและสร้างทางลัด" ของ Chrome
 */
function getActiveFaviconUrl() {
  try {
    const saved = PropertiesService.getScriptProperties().getProperty('CUSTOM_APP_ICON_URL');
    if (saved && saved.indexOf('http') === 0) {
      return saved;
    }
  } catch (e) {}
  return CONFIG.DEFAULT_3D_ICON_URL;
}

/**
 * 2.2 บันทึกรูปไอคอน 3 มิติสวยๆ ลงใน Google Drive และตั้งเป็น Favicon หลักของเว็บแอปอัตโนมัติ
 */
function saveCustomAppIconToDrive(base64DataUrl, directPresetUrl) {
  try {
    if (base64DataUrl && base64DataUrl.indexOf('data:image/png;base64,') === 0) {
      const folder = getOrCreateMainDriveFolder();
      const base64Clean = base64DataUrl.replace(/^data:image\/png;base64,/, '');
      const decoded = Utilities.base64Decode(base64Clean);
      const blob = Utilities.newBlob(decoded, 'image/png', 'MaintPro_3D_App_Icon_512.png');

      // ลบไฟล์ไอคอนเก่าถ้ามีเพื่อไม่ให้เปลืองพื้นที่
      const oldFiles = folder.getFilesByName('MaintPro_3D_App_Icon_512.png');
      while (oldFiles.hasNext()) {
        try { oldFiles.next().setTrashed(true); } catch (delErr) {}
      }

      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      const fileId = file.getId();
      // ใช้ลิงก์ Thumbnail PNG ความละเอียดสูงของ Google Drive เพื่อให้ Chrome Android แสดงรูปไอคอน 3D ทันที
      const publicPngUrl = 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w512&.png';
      PropertiesService.getScriptProperties().setProperty('CUSTOM_APP_ICON_URL', publicPngUrl);

      return {
        status: 'SUCCESS',
        iconUrl: publicPngUrl,
        webAppUrl: getWebAppExecUrl(),
        message: 'บันทึกไอคอน 3 มิติลงระบบสำเร็จ!'
      };
    }

    if (directPresetUrl && directPresetUrl.indexOf('http') === 0) {
      PropertiesService.getScriptProperties().setProperty('CUSTOM_APP_ICON_URL', directPresetUrl);
      return {
        status: 'SUCCESS',
        iconUrl: directPresetUrl,
        webAppUrl: getWebAppExecUrl(),
        message: 'ตั้งค่าไอคอนแอปใหม่สำเร็จ!'
      };
    }

    return { status: 'ERROR', message: 'ข้อมูลรูปไอคอนไม่ถูกต้อง' };
  } catch (err) {
    const fallbackUrl = (directPresetUrl && directPresetUrl.indexOf('http') === 0)
      ? directPresetUrl
      : CONFIG.DEFAULT_3D_ICON_URL;
    PropertiesService.getScriptProperties().setProperty('CUSTOM_APP_ICON_URL', fallbackUrl);
    return {
      status: 'SUCCESS',
      iconUrl: fallbackUrl,
      webAppUrl: getWebAppExecUrl(),
      message: 'ตั้งค่าไอคอน 3 มิติสำเร็จ'
    };
  }
}

/**
 * 2.3 สร้างไฟล์ติดตั้งแอปสำหรับ iOS (.mobileconfig / WebClip เสมือน .ipa) ลงใน Google Drive เพื่อดาวน์โหลดติดตั้งลงหน้าจอ iPhone
 */
function createIosInstallerFileInDrive(targetUrl, base64IconDataUrl) {
  try {
    const folder = getOrCreateMainDriveFolder();
    const cleanUrl = targetUrl || getWebAppExecUrl() || 'https://script.google.com';
    const base64Icon = (base64IconDataUrl || '').replace(/^data:image\/png;base64,/, '');
    const uuid1 = 'A4MAINT-1111-2222-3333-' + new Date().getTime();
    const uuid2 = 'A4MAINT-4444-5555-6666-' + new Date().getTime();

    const xml = '<?xml version="1.0" encoding="UTF-8"?>\n' +
      '<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">\n' +
      '<plist version="1.0">\n<dict>\n' +
      '  <key>PayloadContent</key>\n  <array>\n    <dict>\n' +
      '      <key>FullScreen</key><true/>\n' +
      (base64Icon ? ('      <key>Icon</key><data>' + base64Icon + '</data>\n') : '') +
      '      <key>IsRemovable</key><true/>\n' +
      '      <key>Label</key><string>MAINT PRO</string>\n' +
      '      <key>PayloadDescription</key><string>Install MAINT PRO Fullscreen App</string>\n' +
      '      <key>PayloadDisplayName</key><string>MAINT PRO App</string>\n' +
      '      <key>PayloadIdentifier</key><string>com.maintpro.webclip.' + uuid1 + '</string>\n' +
      '      <key>PayloadType</key><string>com.apple.webClip.managed</string>\n' +
      '      <key>PayloadUUID</key><string>' + uuid1 + '</string>\n' +
      '      <key>PayloadVersion</key><integer>1</integer>\n' +
      '      <key>Precomposed</key><true/>\n' +
      '      <key>URL</key><string>' + cleanUrl + '</string>\n' +
      '    </dict>\n  </array>\n' +
      '  <key>PayloadDisplayName</key><string>MAINT PRO iOS App Installer</string>\n' +
      '  <key>PayloadIdentifier</key><string>com.maintpro.profile.' + uuid2 + '</string>\n' +
      '  <key>PayloadRemovalDisallowed</key><false/>\n' +
      '  <key>PayloadType</key><string>Configuration</string>\n' +
      '  <key>PayloadUUID</key><string>' + uuid2 + '</string>\n' +
      '  <key>PayloadVersion</key><integer>1</integer>\n' +
      '</dict>\n</plist>';

    const oldFiles = folder.getFilesByName('MAINT_PRO_iOS_App.mobileconfig');
    while (oldFiles.hasNext()) {
      try { oldFiles.next().setTrashed(true); } catch (e) {}
    }

    const blob = Utilities.newBlob(xml, 'application/x-apple-aspen-config', 'MAINT_PRO_iOS_App.mobileconfig');
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return {
      status: 'SUCCESS',
      downloadUrl: 'https://drive.google.com/uc?export=download&id=' + file.getId()
    };
  } catch (err) {
    return { status: 'ERROR', message: err.toString() };
  }
}

/**
 * 3. รับข้อมูลใบงานซ่อมบำรุงจากหน้าเว็บแอป (ผ่าน POST)
 */
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return createJsonResponse({ status: 'ERROR', message: 'ไม่พบข้อมูลที่ส่งมา (Empty Payload)' });
    }
    const data = JSON.parse(e.postData.contents);
    if (data.recordType === 'EQUIPMENT_REQUISITION') {
      const eqResult = processSaveEquipmentPayload(data);
      return createJsonResponse(eqResult);
    }
    const result = processSaveReportPayload(data);
    return createJsonResponse(result);
  } catch (err) {
    return createJsonResponse({
      status: 'ERROR',
      message: err.toString()
    });
  }
}

/**
 * 3.1 ฟังก์ชันสำหรับเรียกบันทึกข้อมูลจากหน้าเว็บแอปโดยตรงผ่าน google.script.run
 */
function saveReportFromClient(data) {
  return processSaveReportPayload(data);
}

/**
 * 3.1.1 ฟังก์ชันสำหรับบันทึกใบเบิก-จ่ายอุปกรณ์เข้าทำงานและอะไหล่ช่างผ่าน google.script.run
 */
function saveEquipmentFromClient(data) {
  return processSaveEquipmentPayload(data);
}

/**
 * 3.1.2 ฟังก์ชันดึงข้อมูลประวัติใบเบิก-คืนอุปกรณ์และอะไหล่ช่างทั้งหมดผ่าน google.script.run
 */
function getEquipmentRecordsFromClient() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureEquipmentSheet(ss);
  return {
    status: 'SUCCESS',
    sheetUrl: ss.getUrl(),
    records: getAllEquipmentRecords()
  };
}

/**
 * 3.2 ฟังก์ชันสำหรับดึงข้อมูลประวัติใบงานทั้งหมดผ่าน google.script.run
 */
function getRecordsFromClient() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureMainSheet(ss);
  return {
    status: 'SUCCESS',
    sheetUrl: ss.getUrl(),
    webAppUrl: getWebAppExecUrl(),
    records: getAllMaintenanceRecords()
  };
}

/**
 * 3.2.1 ดึงลิงก์ Web App URL (.../exec) จริงของโปรเจ็กต์ เพื่อให้สแกน QR Code และติดตั้งลงมือถือได้ทันที
 */
function getWebAppExecUrl() {
  try {
    return ScriptApp.getService().getUrl() || '';
  } catch (e) {
    return '';
  }
}

/**
 * 3.3 ประมวลผลการบันทึกข้อมูลใบงานลง Google Sheets และ Google Drive
 */
function processSaveReportPayload(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ensureMainSheet(ss);
  const detailSheet = ensureDetailSheet(ss);

  const now = new Date();
  const recordId = data.id || ('REC-' + Utilities.formatDate(now, 'GMT+7', 'yyyyMMdd-HHmmss'));
  const serviceDate = data.date || Utilities.formatDate(now, 'GMT+7', 'yyyy-MM-dd');
  const machineCode = data.machineCode || 'MCH-000';
  const machineName = data.machineName || '-';
  const status = data.status || 'OPEN';
  const snapshot = data.snapshot || {};
  const techName = snapshot.techName || data.techName || 'ทีมวิศวกรรมซ่อมบำรุง';

  // แยกค่าหัวข้อที่ 6: ความร้อน Temperature (°C) 3 ค่า และ ความสั่นสะเทือน Vibration (mm/s) 3 ค่า
  const tempArr = Array.isArray(snapshot.tempValues) ? snapshot.tempValues : ['-', '-', '-'];
  const vibArr = Array.isArray(snapshot.vibValues) ? snapshot.vibValues : ['-', '-', '-'];

  // แยกค่าหัวข้อที่ 7: Alarm Code และข้อเสนอแนะ Comment
  const alarmCode = data.alarmCode || snapshot.alarmCode || 'NONE';
  const comment = data.comment || snapshot.comment || '-';
  const photoCount = data.photoCount || 0;

  // แยกสรุปรายละเอียดของหัวข้อที่ 1 - 7 ออกจากกันอย่างชัดเจน
  const sectionSummaries = extractSeparatedSectionsSummary(snapshot.pages || []);

  // บันทึกรูปภาพลงโฟลเดอร์ Google Drive อัตโนมัติ (แยกโฟลเดอร์ตามรหัสเครื่องจักรและวันที่)
  let jobDriveFolderUrl = '-';
  if (snapshot.pages && snapshot.pages.length > 0) {
    jobDriveFolderUrl = saveJobPhotosToDriveFolder(recordId, machineCode, serviceDate, snapshot.pages, detailSheet);
  }

  // สร้าง Snapshot แบบประหยัดพื้นที่สำหรับเก็บในเซลล์เพื่อให้โหลดกลับมาตรวจสอบได้
  const lightSnapshot = buildLightweightSnapshot(snapshot);

  // บันทึกแถวข้อมูลหลักลงในตาราง Maintenance_Report_Log
  mainSheet.appendRow([
    now,
    recordId,
    serviceDate,
    machineCode,
    machineName,
    techName,
    status,
    sectionSummaries.SEC_1,
    sectionSummaries.SEC_2,
    sectionSummaries.SEC_3,
    sectionSummaries.SEC_4,
    sectionSummaries.SEC_5,
    tempArr[0] || '-',
    tempArr[1] || '-',
    tempArr[2] || '-',
    vibArr[0] || '-',
    vibArr[1] || '-',
    vibArr[2] || '-',
    alarmCode,
    comment,
    photoCount,
    jobDriveFolderUrl,
    JSON.stringify(lightSnapshot)
  ]);

  return {
    status: 'SUCCESS',
    message: 'บันทึกข้อมูลใบงานซ่อมบำรุงเข้า Google Sheets และ Google Drive สำเร็จ',
    recordId: recordId,
    driveFolderUrl: jobDriveFolderUrl,
    sheetUrl: ss.getUrl(),
    records: getAllMaintenanceRecords()
  };
}

/**
 * 4. แยกสรุปรายละเอียดของหัวข้อที่ 1 - 7 ออกจากกันอย่างชัดเจน
 */
function extractSeparatedSectionsSummary(pages) {
  const result = {
    SEC_1: '-',
    SEC_2: '-',
    SEC_3: '-',
    SEC_4: '-',
    SEC_5: '-',
    SEC_6: '-',
    SEC_7: '-'
  };

  if (!Array.isArray(pages)) return result;

  const bucket = { SEC_1: [], SEC_2: [], SEC_3: [], SEC_4: [], SEC_5: [], SEC_6: [], SEC_7: [] };

  pages.forEach((page, pIdx) => {
    const code = page.sectionCode;
    if (!bucket[code] || !Array.isArray(page.slots)) return;

    page.slots.forEach((slot) => {
      const hasImg = slot.imageSrc ? '📷มีรูป' : '⚪ไม่มีรูป';
      const status = slot.statusTag || '';
      const sub = slot.subTitle || '';
      const remarks = Array.isArray(slot.remarks) ? slot.remarks.join(', ') : '';
      bucket[code].push(`[${ status } | ${ sub }: ${ remarks } (${ hasImg })]`);
    });
  });

  Object.keys(bucket).forEach((key) => {
    if (bucket[key].length > 0) {
      result[key] = bucket[key].join('\n');
    }
  });

  return result;
}

/**
 * 5. บันทึกรูปภาพ Base64 จากหน้ากระดาษ A4 ลงในโฟลเดอร์ Google Drive แยกตามใบงาน
 */
function saveJobPhotosToDriveFolder(recordId, machineCode, serviceDate, pages, detailSheet) {
  try {
    const mainFolder = getOrCreateMainDriveFolder();
    const subFolderName = `${serviceDate}_${machineCode}_${recordId}`;
    const jobFolder = mainFolder.createFolder(subFolderName);

    pages.forEach((page, pageIdx) => {
      const secCode = page.sectionCode || 'SEC_1';
      if (!Array.isArray(page.slots)) return;

      page.slots.forEach((slot, slotIdx) => {
        const remarksText = Array.isArray(slot.remarks) ? slot.remarks.join(' | ') : '';
        let fileUrl = '-';

        // บันทึกเฉพาะรูปที่ช่างอัปโหลดหรือถ่ายจริง (ข้าม SVG ตัวอย่าง)
        if (slot.imageSrc && slot.imageSrc.indexOf('data:image/jpeg;base64,') === 0) {
          const base64Data = slot.imageSrc.replace(/^data:image\/\w+;base64,/, '');
          const decoded = Utilities.base64Decode(base64Data);
          const safeStatus = (slot.statusTag || 'Photo').replace(/[^a-zA-Z0-9ก-๙_-]/g, '_');
          const fileName = `P${pageIdx + 1}_Slot${slotIdx + 1}_${secCode}_${safeStatus}.jpg`;
          const blob = Utilities.newBlob(decoded, 'image/jpeg', fileName);
          const file = jobFolder.createFile(blob);
          fileUrl = file.getUrl();
        }

        // บันทึกรายการรูปและหมายเหตุใต้รูปลงตารางย่อย (Sections_1_to_7_Photos)
        detailSheet.appendRow([
          new Date(),
          recordId,
          serviceDate,
          machineCode,
          `หน้าที่ ${pageIdx + 1}`,
          secCode,
          slot.sectionTitle || '-',
          slot.statusTag || '-',
          slot.subTitle || '-',
          remarksText,
          fileUrl
        ]);
      });
    });

    return jobFolder.getUrl();
  } catch (e) {
    Logger.log('Drive Save Warning: ' + e.toString());
    return '-';
  }
}

/**
 * 6. ดึงข้อมูลประวัติใบงานทั้งหมดจาก Google Sheets ส่งกลับไปแสดงในตารางตรวจสอบของเว็บแอป
 */
function getAllMaintenanceRecords() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.MAIN_SHEET_NAME);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  const data = sheet.getDataRange().getValues();
  const records = [];

  // อ่านจากแถวล่าสุดขึ้นบน (ย้อนหลังสูงสุด 100 รายการ)
  const startRow = data.length - 1;
  const minRow = Math.max(1, data.length - 100);

  for (let i = startRow; i >= minRow; i--) {
    const row = data[i];
    let parsedSnapshot = null;
    try {
      if (row[22]) parsedSnapshot = JSON.parse(row[22]);
    } catch (err) {
      parsedSnapshot = null;
    }

    const formattedDate = row[2] instanceof Date
      ? Utilities.formatDate(row[2], 'GMT+7', 'yyyy-MM-dd')
      : String(row[2] || '');

    records.push({
      id: String(row[1] || ''),
      date: formattedDate,
      machineCode: String(row[3] || ''),
      machineName: String(row[4] || ''),
      techName: String(row[5] || ''),
      status: String(row[6] || 'OPEN'),
      sec1Summary: String(row[7] || '-'),
      sec2Summary: String(row[8] || '-'),
      sec3Summary: String(row[9] || '-'),
      sec4Summary: String(row[10] || '-'),
      sec5Summary: String(row[11] || '-'),
      temps: `${row[12] || '-'} / ${row[13] || '-'} / ${row[14] || '-'} °C`,
      vibs: `${row[15] || '-'} / ${row[16] || '-'} / ${row[17] || '-'} mm/s`,
      alarmCode: String(row[18] || 'NONE'),
      comment: String(row[19] || '-'),
      photoCount: Number(row[20] || 0),
      driveFolderUrl: String(row[21] || '-'),
      snapshot: parsedSnapshot
    });
  }

  return records;
}

/**
 * 7. สร้างและตกแต่งหัวตารางหลัก (Maintenance_Report_Log) อัตโนมัติหากยังไม่มี
 */
function ensureMainSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.MAIN_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.MAIN_SHEET_NAME);
    const headers = [
      'วัน-เวลาที่บันทึก',
      'เลขที่ใบงาน (Record ID)',
      'วันที่เข้าทำ (Service Date)',
      'รหัสเครื่องจักร (Machine Code)',
      'ชื่อเครื่องจักร / แผนก',
      'ผู้ปฏิบัติงาน / ช่างเทคนิค',
      'สถานะใบงาน',
      '1. OVERHAUL MOTOR (ก่อนทำ-หลังทำ)',
      '2. OVERHAUL BLOWER (ก่อนทำ-หลังทำ)',
      '3. เปลี่ยนสายพาน & Alignment Laser (ก่อนทำ-หลังทำ)',
      '4. CLEANING (ก่อนทำ-หลังทำ)',
      '5. PM (Preventive Maintenance)',
      '6.1 Temp จุดที่ 1 (°C)',
      '6.1 Temp จุดที่ 2 (°C)',
      '6.1 Temp จุดที่ 3 (°C)',
      '6.2 Vibration แกน H (mm/s)',
      '6.2 Vibration แกน V (mm/s)',
      '6.2 Vibration แกน A (mm/s)',
      '7.1 รหัสแจ้งเตือน (Alarm Code)',
      '7.2 ข้อเสนอแนะ (Comment)',
      'จำนวนรูปภาพ',
      'ลิงก์โฟลเดอร์รูป (Google Drive)',
      'JSON_Snapshot'
    ];
    sheet.appendRow(headers);

    // ตกแต่งสีหัวตารางให้ดูเป็นมืออาชีพ
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#0f172a');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * 8.1 สร้างตารางหลักสำหรับระบบเบิก-จ่าย ขึ้นของ และรับคืนอุปกรณ์ซ่อมบำรุง / อะไหล่ช่าง
 */
function ensureEquipmentSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.EQUIP_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.EQUIP_SHEET_NAME);
    const headers = [
      'วัน-เวลาที่บันทึก',
      'เลขที่ใบเบิก (Doc ID)',
      'วันที่เบิก / ขึ้นของ',
      'วันที่นำมาคืน',
      'สถานะใบเบิก / อุปกรณ์',
      'รหัสเครื่องจักร / หน้างาน',
      'ผู้เบิกจ่าย / นำของขึ้นเข้าทำงาน',
      'ผู้จ่ายของ (Store/승인)',
      'ผู้นำของมาคืน (เมื่อจบงาน)',
      'ผู้ตรวจรับคืน',
      'จำนวนรายการทั้งหมด',
      'สรุปรายการอุปกรณ์ซ่อมบำรุง & อะไหล่ช่าง',
      'หมายเหตุการเบิก-คืน',
      'JSON_Equip_Snapshot'
    ];
    sheet.appendRow(headers);

    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#065f46');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * 8.2 สร้างตารางย่อยเก็บรายการอุปกรณ์ซ่อมบำรุงและอะไหล่ช่างรายชิ้น
 */
function ensureEquipmentItemsSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.EQUIP_ITEMS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.EQUIP_ITEMS_SHEET_NAME);
    const headers = [
      'วัน-เวลาที่บันทึก',
      'เลขที่ใบเบิก (Doc ID)',
      'วันที่เบิก / ขึ้นของ',
      'วันที่นำมาคืน',
      'ผู้เบิกจ่าย',
      'ผู้นำของมาคืน',
      'รหัสอุปกรณ์/อะไหล่',
      'ชื่ออุปกรณ์ซ่อมบำรุง / อะไหล่ช่าง',
      'หมวดหมู่ (อุปกรณ์ซ่อม/อะไหล่)',
      'จำนวนที่เบิก',
      'จำนวนที่คืน',
      'หน่วยนับ',
      'สถานะของอุปกรณ์ซ่อม',
      'หมายเหตุรายชิ้น'
    ];
    sheet.appendRow(headers);

    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#047857');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * 8.3 บันทึกหรืออัปเดตข้อมูลใบเบิก-คืนอุปกรณ์ซ่อมบำรุงและอะไหล่ช่างลง Google Sheets
 */
function processSaveEquipmentPayload(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ensureEquipmentSheet(ss);
  const itemsSheet = ensureEquipmentItemsSheet(ss);

  const now = new Date();
  const docId = data.docId || ('EQ-' + Utilities.formatDate(now, 'GMT+7', 'yyyyMMdd-HHmmss'));
  const checkoutDate = data.checkoutDate || Utilities.formatDate(now, 'GMT+7', 'yyyy-MM-dd');
  const returnDate = data.returnDate || '-';
  const overallStatus = data.overallStatus || 'กำลังเบิกขึ้นของเข้าทำงาน';
  const machineOrSite = data.machineOrSite || '-';
  const requesterName = data.requesterName || '-';
  const issuerName = data.issuerName || '-';
  const returnerName = data.returnerName || '-';
  const receiverName = data.receiverName || '-';
  const note = data.note || '-';
  const items = Array.isArray(data.items) ? data.items : [];

  const summaryLines = items.map(function(it, idx) {
    return (idx + 1) + '. [' + (it.category || 'อุปกรณ์ช่าง') + '] ' + (it.name || '-') +
      ' (เบิก: ' + (it.qtyOut || 0) + ' / คืน: ' + (it.qtyReturn || 0) + ' ' + (it.unit || 'ชิ้น') +
      ' | สถานะ: ' + (it.itemStatus || '-') + ')';
  }).join('\n');

  // ตรวจสอบว่ามีเลขที่ใบเบิกนี้อยู่แล้วหรือไม่ ถ้ามีให้อัปเดตแถวเดิม (เช่น ตอนกดรับของคืน)
  const allValues = mainSheet.getDataRange().getValues();
  let foundRowIndex = -1;
  for (let r = 1; r < allValues.length; r++) {
    if (String(allValues[r][1]) === String(docId)) {
      foundRowIndex = r + 1;
      break;
    }
  }

  const rowValues = [
    now,
    docId,
    checkoutDate,
    returnDate,
    overallStatus,
    machineOrSite,
    requesterName,
    issuerName,
    returnerName,
    receiverName,
    items.length,
    summaryLines || '-',
    note,
    JSON.stringify(buildLightweightEquipSnapshot(data))
  ];

  if (foundRowIndex !== -1) {
    mainSheet.getRange(foundRowIndex, 1, 1, rowValues.length).setValues([rowValues]);
  } else {
    mainSheet.appendRow(rowValues);
  }

  // บันทึกรายการย่อยลงตาราง Equipment_Items_Detail
  items.forEach(function(it) {
    itemsSheet.appendRow([
      now,
      docId,
      checkoutDate,
      returnDate,
      requesterName,
      returnerName,
      it.code || '-',
      it.name || '-',
      it.category || 'อุปกรณ์ซ่อมบำรุง',
      it.qtyOut || 1,
      it.qtyReturn || 0,
      it.unit || 'ชุด',
      it.itemStatus || 'กำลังใช้งานหน้างาน',
      it.remark || '-'
    ]);
  });

  return {
    status: 'SUCCESS',
    message: 'บันทึกใบเบิก-จ่าย รูปถ่ายหลักฐาน และสถานะการคืนอุปกรณ์ลง Google Sheets สำเร็จ',
    docId: docId,
    sheetUrl: ss.getUrl(),
    records: getAllEquipmentRecords()
  };
}

/**
 * 8.3.1 ย่อขนาดข้อมูลใบเบิก-คืนอุปกรณ์เพื่อให้เก็บรูปตัวอย่างและดึงกลับมาดูได้ใน Google Sheets
 */
function buildLightweightEquipSnapshot(data) {
  if (!data) return null;
  const copy = JSON.parse(JSON.stringify(data));
  // เก็บรูปหลักฐานเบิก-คืนขนาดย่อไม่ให้เกินลิมิต 50,000 ตัวอักษรของเซลล์ Google Sheets
  if (copy.checkoutPhoto && copy.checkoutPhoto.length > 18000) {
    copy.checkoutPhoto = copy.checkoutPhoto;
  }
  if (copy.returnPhoto && copy.returnPhoto.length > 18000) {
    copy.returnPhoto = copy.returnPhoto;
  }
  if (Array.isArray(copy.items)) {
    copy.items.forEach(function(it) {
      if (it.photo && it.photo.length > 8000) {
        it.photo = null;
      }
    });
  }
  return copy;
}

/**
 * 8.4 ดึงประวัติการเบิก-จ่ายและรับคืนอุปกรณ์ซ่อมบำรุงทั้งหมดจาก Google Sheets
 */
function getAllEquipmentRecords() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.EQUIP_SHEET_NAME);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  const data = sheet.getDataRange().getValues();
  const records = [];
  const startRow = data.length - 1;
  const minRow = Math.max(1, data.length - 150);

  for (let i = startRow; i >= minRow; i--) {
    const row = data[i];
    let parsedSnapshot = null;
    try {
      if (row[13]) parsedSnapshot = JSON.parse(row[13]);
    } catch (e) {
      parsedSnapshot = null;
    }

    const fmtOut = row[2] instanceof Date
      ? Utilities.formatDate(row[2], 'GMT+7', 'yyyy-MM-dd')
      : String(row[2] || '-');
    const fmtRet = row[3] instanceof Date
      ? Utilities.formatDate(row[3], 'GMT+7', 'yyyy-MM-dd')
      : String(row[3] || '-');

    if (parsedSnapshot) {
      parsedSnapshot.checkoutDate = fmtOut;
      parsedSnapshot.returnDate = fmtRet;
      records.push(parsedSnapshot);
    } else {
      records.push({
        docId: String(row[1] || ''),
        checkoutDate: fmtOut,
        returnDate: fmtRet,
        overallStatus: String(row[4] || ''),
        machineOrSite: String(row[5] || ''),
        requesterName: String(row[6] || ''),
        issuerName: String(row[7] || ''),
        returnerName: String(row[8] || ''),
        receiverName: String(row[9] || ''),
        itemCount: Number(row[10] || 0),
        summaryText: String(row[11] || ''),
        note: String(row[12] || ''),
        items: []
      });
    }
  }
  return records;
}

/**
 * 9. จัดการโฟลเดอร์หลักใน Google Drive
 */
function getOrCreateMainDriveFolder() {
  const folders = DriveApp.getFoldersByName(CONFIG.DRIVE_MAIN_FOLDER_NAME);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(CONFIG.DRIVE_MAIN_FOLDER_NAME);
}

/**
 * 10. ย่อขนาด Snapshot เพื่อไม่ให้เกินลิมิต 50,000 ตัวอักษรต่อเซลล์ของ Google Sheets
 */
function buildLightweightSnapshot(snapshot) {
  if (!snapshot) return null;
  const copy = JSON.parse(JSON.stringify(snapshot));
  if (Array.isArray(copy.pages)) {
    copy.pages.forEach((p) => {
      if (Array.isArray(p.slots)) {
        p.slots.forEach((s) => {
          // เก็บเฉพาะข้อมูลโครงสร้างและข้อความหมายเหตุใต้รูป เพื่อให้โหลดกลับมาตรวจได้รวดเร็ว
          if (s.imageSrc && s.imageSrc.length > 2000) {
            s.imageSrc = null;
          }
        });
      }
    });
  }
  return copy;
}

/**
 * 11. ส่งค่ากลับในรูปแบบ JSON มาตรฐาน
 */
function createJsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 12. โค้ดหน้าเว็บแอป A4 Maintenance Report Pro แบบ Self-Contained ใน Code.gs
 * ทำให้วางเฉพาะไฟล์ Code.gs ไฟล์เดียวแล้วกด Deploy เป็น Web App เปิดใช้งานได้ทันที
 */
function getWebAppHtml() {
  return [
    '<!DOCTYPE html>',
    '<html lang="th">',
    '<head>',
    '<meta charset="UTF-8">',
    '<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=5.0">',
    '<title>ระบบรายงานซ่อมบำรุงเครื่องจักร A4 Pro (Google Sheets Web App)</title>',
    '<script src="https://cdn.tailwindcss.com"><\/script>',
    '<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.1/css/all.min.css">',
    '<link href="https://fonts.googleapis.com/css2?family=Prompt:wght@400;500;600;700&family=Sarabun:wght@400;500;600;700&display=swap" rel="stylesheet">',
    '<script src="https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js"><\/script>',
    '<style id="dynamicPageStyle">@page { size: A4 portrait; margin: 0; }</style>',
    '<style>',
    ':root { --a4-width: 210mm; --a4-height: 297mm; --frame-size: 6cm; }',
    'body { font-family: "Prompt", "Sarabun", sans-serif; background:#0f172a; color:#1e293b; overflow-x:hidden; }',
    '.a4-sheet { font-family:"Sarabun",sans-serif; width:var(--a4-width); height:var(--a4-height); min-width:var(--a4-width); min-height:var(--a4-height); max-width:var(--a4-width); max-height:var(--a4-height); background:#fff; box-shadow:0 20px 50px rgba(0,0,0,0.45); margin:0 auto 24px auto; padding:8mm 10mm; box-sizing:border-box; position:relative; overflow:hidden; display:flex; flex-direction:column; justify-content:space-between; page-break-after:always; break-after:page; }',
    '.a4-sheet.landscape { --a4-width:297mm; --a4-height:210mm; width:297mm; height:210mm; min-width:297mm; min-height:210mm; max-width:297mm; max-height:210mm; }',
    '.photo-frame-box { width:var(--frame-size); height:var(--frame-size); max-width:100%; max-height:100%; margin:0 auto; position:relative; display:flex; align-items:center; justify-content:center; overflow:hidden; background:#f8fafc; border:1.5px dashed #cbd5e1; border-radius:6px; }',
    '.photo-frame-box.has-image { border:1.5px solid #334155; background:#fff; }',
    '.photo-frame-box img { width:100%; height:100%; object-position:center center; display:block; margin:auto; }',
    '#a4PreviewViewport { width:100%; overflow-x:auto; display:flex; flex-direction:column; align-items:center; padding:16px 8px 90px 8px; }',
    '.a4-scale-wrapper { transform-origin:top center; display:flex; flex-direction:column; align-items:center; }',
    '@media print { body{background:#fff!important;margin:0!important;padding:0!important;} .no-print{display:none!important;} #a4PreviewViewport{padding:0!important;margin:0!important;overflow:visible!important;display:block!important;} .a4-scale-wrapper{transform:none!important;display:block!important;} .a4-sheet{margin:0!important;box-shadow:none!important;border:none!important;-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important;} .slot-action-overlay{display:none!important;} }',
    '</style>',
    '</head>',
    '<body class="min-h-screen flex flex-col bg-slate-900 text-slate-100">',
    '<header class="no-print sticky top-0 z-40 bg-slate-900/95 backdrop-blur border-b border-slate-700 px-3 py-2 shadow-lg">',
    '  <div class="max-w-[1600px] mx-auto flex flex-wrap items-center justify-between gap-2">',
    '      <div class="flex items-center gap-3">',
    '        <div onclick="cycleHeaderAppIcon()" class="relative group cursor-pointer shrink-0" title="คลิกเพื่อเปลี่ยนสไตล์ไอคอนแอป 3 มิติ">',
    '          <div class="absolute -inset-0.5 bg-gradient-to-r from-cyan-400 via-amber-400 to-emerald-400 rounded-2xl blur opacity-70 group-hover:opacity-100 transition"></div>',
    '          <img id="headerBrandAppIcon" class="relative w-11 h-11 rounded-2xl shadow-xl border-2 border-white/80 object-cover" alt="App Icon">',
    '        </div>',
    '        <div>',
    '          <div class="flex items-center gap-2">',
    '            <h1 class="font-extrabold text-sm sm:text-base bg-gradient-to-r from-white via-cyan-200 to-emerald-300 bg-clip-text text-transparent">MAINTENANCE & EQUIPMENT PRO</h1>',
    '            <span id="jobStatusBadge" class="px-2 py-0.5 text-[11px] font-semibold rounded-full bg-emerald-500/20 text-emerald-300 border border-emerald-500/40">เปิดงาน (OPEN)</span>',
    '          </div>',
    '          <p class="text-[11px] text-slate-400 hidden sm:block">ระบบเบิก-จ่ายอุปกรณ์เข้าทำงาน ถ่ายรูปหลักฐานเบิก-คืน & ใบรายงานซ่อมบำรุง A4</p>',
    '        </div>',
    '      </div>',
    '      <div class="flex bg-slate-800 p-1 rounded-xl border border-slate-700 gap-1">',
    '        <button onclick="switchMainAppMode(\'A4_REPORT\')" id="tabBtnReport" class="px-3 py-1.5 rounded-lg text-xs font-bold bg-cyan-500 text-slate-950"><i class="fa-solid fa-file-invoice mr-1"></i>1. ใบรายงานซ่อมบำรุง A4 (หัวข้อ 1-7)</button>',
    '        <button onclick="switchMainAppMode(\'EQUIPMENT\')" id="tabBtnEquip" class="px-3 py-1.5 rounded-lg text-xs font-bold text-slate-300 hover:bg-slate-700"><i class="fa-solid fa-toolbox mr-1"></i>2. เบิก-จ่าย / คืนอุปกรณ์ซ่อม & อะไหล่ช่าง</button>',
    '      </div>',
    '      <div class="flex flex-wrap items-center gap-1.5">',
    '        <div id="topReportActionBtns" class="flex flex-wrap items-center gap-1.5">',
    '          <button onclick="saveToGoogleSheetsDirect()" class="px-3 py-1.5 rounded-lg text-xs font-extrabold bg-amber-400 hover:bg-amber-300 text-slate-950 shadow"><i class="fa-solid fa-floppy-disk mr-1"></i>บันทึก (รอตรวจสอบ)</button>',
    '          <button onclick="toggleJobStatus()" id="btnJobToggle" class="px-2.5 py-1.5 rounded-lg text-xs font-semibold bg-slate-700 hover:bg-slate-600 text-white"><i class="fa-solid fa-clipboard-check mr-1"></i><span id="btnJobToggleText">ปิดงาน</span></button>',
    '          <button onclick="triggerBatchUpload()" class="px-2.5 py-1.5 rounded-lg text-xs font-semibold bg-blue-600 hover:bg-blue-500 text-white"><i class="fa-solid fa-images mr-1"></i>เพิ่มหลายรูป</button>',
    '          <input type="file" id="batchFileInput" accept="image/*" multiple class="hidden" onchange="handleBatchFiles(event)">',
    '          <button onclick="openSheetsModal()" class="px-2.5 py-1.5 rounded-lg text-xs font-semibold bg-emerald-600 hover:bg-emerald-500 text-white"><i class="fa-solid fa-clock-rotate-left mr-1"></i>ประวัติใบงาน</button>',
    '          <button onclick="exportToPDF()" class="px-2.5 py-1.5 rounded-lg text-xs font-semibold bg-rose-600 hover:bg-rose-500 text-white"><i class="fa-solid fa-file-pdf mr-1"></i>.PDF</button>',
    '          <button onclick="exportToWord()" class="px-2.5 py-1.5 rounded-lg text-xs font-semibold bg-sky-600 hover:bg-sky-500 text-white"><i class="fa-solid fa-file-word mr-1"></i>.WORD</button>',
    '          <button onclick="window.print()" class="px-2.5 py-1.5 rounded-lg text-xs font-semibold bg-slate-700 hover:bg-slate-600 text-white"><i class="fa-solid fa-print mr-1"></i>ปริ้น A4</button>',
    '        </div>',
    '      </div>',
    '  </div>',
    '</header>',
    '<!-- ==================== EQUIPMENT & SPARE PARTS CHECK-OUT / RETURN WORKSPACE ==================== -->',
    '<section id="equipmentAppWorkspace" class="hidden flex-1 max-w-[1600px] w-full mx-auto p-3 sm:p-4 space-y-4 overflow-y-auto">',
    '  <div class="grid grid-cols-2 sm:grid-cols-4 gap-2.5">',
    '    <div onclick="openEqStatusModal(\'ALL\')" class="cursor-pointer hover:border-cyan-400 transition bg-slate-800/90 border-2 border-slate-700 rounded-xl p-3 flex items-center justify-between"><div><div class="text-[11px] text-slate-300 font-bold">ใบเบิก-จ่ายทั้งหมด <span class="text-[9px] text-cyan-300 underline">(กดดูสถานะ)</span></div><div id="eqStatTotal" class="text-xl font-extrabold text-white">0</div></div><i class="fa-solid fa-clipboard-list text-2xl text-cyan-400"></i></div>',
    '    <div onclick="openEqStatusModal(\'ACTIVE_OUT\')" class="cursor-pointer hover:border-amber-400 transition bg-slate-800/90 border-2 border-amber-500/40 rounded-xl p-3 flex items-center justify-between"><div><div class="text-[11px] text-amber-300 font-bold">กำลังเบิกขึ้นของเข้าทำงาน (ยังไม่คืน)</div><div id="eqStatOut" class="text-xl font-extrabold text-amber-400">0</div></div><i class="fa-solid fa-truck-ramp-box text-2xl text-amber-400"></i></div>',
    '    <div onclick="openEqStatusModal(\'RETURNED\')" class="cursor-pointer hover:border-emerald-400 transition bg-slate-800/90 border-2 border-emerald-500/40 rounded-xl p-3 flex items-center justify-between"><div><div class="text-[11px] text-emerald-300 font-bold">นำของมาคืนครบแล้ว</div><div id="eqStatReturned" class="text-xl font-extrabold text-emerald-400">0</div></div><i class="fa-solid fa-circle-check text-2xl text-emerald-400"></i></div>',
    '    <div onclick="openEqStatusModal(\'REPAIR\')" class="cursor-pointer hover:border-rose-400 transition bg-slate-800/90 border-2 border-rose-500/40 rounded-xl p-3 flex items-center justify-between"><div><div class="text-[11px] text-rose-300 font-bold">อุปกรณ์ซ่อมชำรุด / รอซ่อม</div><div id="eqStatRepair" class="text-xl font-extrabold text-rose-400">0</div></div><i class="fa-solid fa-screwdriver-wrench text-2xl text-rose-400"></i></div>',
    '  </div>',
    '  <div class="grid grid-cols-1 lg:grid-cols-12 gap-4">',
    '    <div class="lg:col-span-7 bg-slate-800/90 border border-slate-700 rounded-2xl p-4 space-y-3">',
    '      <div class="flex flex-wrap justify-between items-center gap-2 border-b border-slate-700 pb-2.5">',
    '        <div><h2 class="text-sm font-bold text-emerald-400"><i class="fa-solid fa-boxes-packing mr-1.5"></i>ฟอร์มเบิก-จ่ายอุปกรณ์เข้าทำงาน ถ่ายรูปหลักฐาน & รับคืนของเมื่อคืน/วันนี้</h2><p class="text-[11px] text-slate-400">ถ่ายรูปตอนเบิกขึ้นของและตอนนำของมาคืน พร้อมประทับวันที่เบิกจ่ายและวันที่คืนบนภาพอัตโนมัติ</p></div>',
    '        <div class="flex gap-1.5"><button onclick="resetEquipForm()" class="px-2.5 py-1 rounded-lg bg-slate-700 hover:bg-slate-600 text-xs text-white">+ สร้างใบเบิกใหม่</button><button onclick="printEquipSlipWithPhotos()" class="px-2.5 py-1 rounded-lg bg-indigo-600 hover:bg-indigo-500 text-xs text-white font-semibold"><i class="fa-solid fa-print mr-1"></i>พิมพ์ใบเบิก A4</button><button onclick="saveEquipRecord()" class="px-3 py-1 rounded-lg bg-emerald-500 hover:bg-emerald-400 text-slate-950 text-xs font-extrabold shadow"><i class="fa-solid fa-floppy-disk mr-1"></i>บันทึกใบเบิก/รับคืน</button></div>',
    '      </div>',
    '      <div class="grid grid-cols-1 sm:grid-cols-4 gap-2.5 bg-slate-900/80 p-3 rounded-xl border border-slate-700">',
    '        <div><label class="block text-[11px] text-slate-300">เลขที่ใบเบิก (Doc ID)</label><input type="text" id="eqDocId" class="w-full bg-slate-800 border border-slate-600 rounded-lg px-2.5 py-1.5 text-xs text-cyan-300 font-bold"></div>',
    '        <div><div class="flex justify-between items-center mb-0.5"><label class="text-[11px] text-amber-300 font-bold">วันที่เบิกจ่าย / ขึ้นของ *</label><button onclick="setQuickEqDate(\'eqCheckoutDate\',0)" class="text-[9px] bg-amber-500/20 text-amber-300 px-1.5 rounded">วันนี้</button></div><input type="date" id="eqCheckoutDate" class="w-full bg-slate-800 border border-amber-500/60 rounded-lg px-2.5 py-1.5 text-xs text-white font-bold"></div>',
    '        <div><div class="flex justify-between items-center mb-0.5"><label class="text-[11px] text-emerald-300 font-bold">วันที่นำมาคืน (ของเมื่อคืน/วันนี้)</label><div class="flex gap-1"><button onclick="setQuickEqDate(\'eqReturnDate\',-1)" class="text-[9px] bg-purple-500/30 text-purple-200 px-1.5 rounded">ของเมื่อคืน</button><button onclick="setQuickEqDate(\'eqReturnDate\',0)" class="text-[9px] bg-emerald-500/20 text-emerald-300 px-1.5 rounded">วันนี้</button></div></div><input type="date" id="eqReturnDate" class="w-full bg-slate-800 border border-emerald-500/60 rounded-lg px-2.5 py-1.5 text-xs text-white font-bold"></div>',
    '        <div><label class="block text-[11px] text-slate-300 font-semibold">สถานะใบเบิก / อุปกรณ์</label><select id="eqOverallStatus" class="w-full bg-slate-800 border border-slate-600 rounded-lg px-2 py-1.5 text-xs text-white font-bold"><option value="รอตรวจสอบ">รอตรวจสอบ (Pending Review)</option><option value="กำลังเบิกขึ้นของเข้าทำงาน">กำลังเบิกขึ้นของเข้าทำงาน</option><option value="นำมาคืนครบแล้ว">นำมาคืนครบแล้ว</option><option value="คืนของเมื่อคืนเรียบร้อย">คืนของเมื่อคืนเรียบร้อย</option><option value="คืนบางส่วน / ค้างคืน">คืนบางส่วน / ค้างคืน</option><option value="อะไหล่ใช้เปลี่ยนหน้างาน">อะไหล่ใช้เปลี่ยนหน้างาน</option><option value="อุปกรณ์ชำรุด / ส่งซ่อม">อุปกรณ์ชำรุด / ส่งซ่อม</option></select></div>',
    '      </div>',
    '      <div class="grid grid-cols-2 sm:grid-cols-4 gap-2.5">',
    '        <div><label class="block text-[11px] text-slate-300">ผู้เบิกจ่าย / นำของขึ้น *</label><input type="text" id="eqRequester" placeholder="ชื่อช่างผู้เบิกของ" class="w-full bg-slate-900 border border-slate-600 rounded-lg px-2.5 py-1.5 text-xs text-white font-semibold"></div>',
    '        <div><label class="block text-[11px] text-slate-300">ผู้จ่ายของ (สโตร์/หัวหน้า)</label><input type="text" id="eqIssuer" placeholder="ชื่อผู้จ่ายของ" class="w-full bg-slate-900 border border-slate-600 rounded-lg px-2.5 py-1.5 text-xs text-white"></div>',
    '        <div><label class="block text-[11px] text-emerald-300">ผู้นำของมาคืน (ของเมื่อคืน/จบงาน)</label><input type="text" id="eqReturner" placeholder="ชื่อผู้นำของมาคืน" class="w-full bg-slate-900 border border-slate-600 rounded-lg px-2.5 py-1.5 text-xs text-white font-semibold"></div>',
    '        <div><label class="block text-[11px] text-slate-300">รหัสเครื่องจักร / หน้างาน</label><input type="text" id="eqMachineSite" placeholder="เช่น MCH-BLW-014" class="w-full bg-slate-900 border border-slate-600 rounded-lg px-2.5 py-1.5 text-xs text-white"></div>',
    '      </div>',
    '      <!-- กล่องถ่ายรูปหลักฐาน วันที่เบิกจ่ายขึ้นของ & วันที่นำของมาคืน -->',
    '      <div class="grid grid-cols-1 sm:grid-cols-2 gap-3 bg-slate-900/95 border border-slate-700 rounded-xl p-3">',
    '        <div class="border border-amber-500/40 bg-slate-800/70 rounded-xl p-2.5 flex items-center gap-3">',
    '          <div onclick="openEquipPhotoCapture(\'EQ_CHECKOUT\')" class="w-20 h-20 rounded-lg bg-slate-900 border-2 border-dashed border-amber-400/60 flex items-center justify-center overflow-hidden cursor-pointer shrink-0">',
    '            <img id="eqCheckoutPhotoPreview" class="hidden w-full h-full object-cover">',
    '            <div id="eqCheckoutPhotoEmpty" class="text-center text-[10px] text-amber-300 p-1"><i class="fa-solid fa-camera text-base block mb-0.5"></i>ถ่ายรูปเบิกของ</div>',
    '          </div>',
    '          <div class="flex-1 space-y-1">',
    '            <div class="text-xs font-bold text-amber-300">1. รูปถ่ายตอนเบิกจ่าย / ขึ้นของ</div>',
    '            <p class="text-[10px] text-slate-400">ประทับวันที่เบิกจ่ายและชื่อผู้เบิกบนภาพอัตโนมัติ</p>',
    '            <div class="flex gap-1.5"><button onclick="openEquipPhotoCapture(\'EQ_CHECKOUT\')" class="px-2.5 py-1 rounded bg-amber-500 hover:bg-amber-400 text-slate-950 text-[11px] font-bold"><i class="fa-solid fa-camera mr-1"></i>กดถ่ายรูปตอนเบิก</button><button onclick="clearEquipMainPhoto(\'EQ_CHECKOUT\')" class="px-2 py-1 rounded bg-slate-700 text-slate-300 text-[10px]">ลบรูป</button></div>',
    '          </div>',
    '        </div>',
    '        <div class="border border-emerald-500/40 bg-slate-800/70 rounded-xl p-2.5 flex items-center gap-3">',
    '          <div onclick="openEquipPhotoCapture(\'EQ_RETURN\')" class="w-20 h-20 rounded-lg bg-slate-900 border-2 border-dashed border-emerald-400/60 flex items-center justify-center overflow-hidden cursor-pointer shrink-0">',
    '            <img id="eqReturnPhotoPreview" class="hidden w-full h-full object-cover">',
    '            <div id="eqReturnPhotoEmpty" class="text-center text-[10px] text-emerald-300 p-1"><i class="fa-solid fa-camera text-base block mb-0.5"></i>ถ่ายรูปคืนของ</div>',
    '          </div>',
    '          <div class="flex-1 space-y-1">',
    '            <div class="text-xs font-bold text-emerald-300">2. รูปถ่ายของเมื่อคืน / ตอนนำมาคืน</div>',
    '            <p class="text-[10px] text-slate-400">ประทับวันที่นำของมาคืนและชื่อผู้คืนบนภาพอัตโนมัติ</p>',
    '            <div class="flex gap-1.5"><button onclick="openEquipPhotoCapture(\'EQ_RETURN\')" class="px-2.5 py-1 rounded bg-emerald-500 hover:bg-emerald-400 text-slate-950 text-[11px] font-bold"><i class="fa-solid fa-camera mr-1"></i>กดถ่ายรูปตอนคืน</button><button onclick="clearEquipMainPhoto(\'EQ_RETURN\')" class="px-2 py-1 rounded bg-slate-700 text-slate-300 text-[10px]">ลบรูป</button></div>',
    '          </div>',
    '        </div>',
    '      </div>',
    '      <div class="bg-slate-900/90 border border-slate-700 rounded-xl p-3 space-y-2">',
    '        <div class="flex flex-wrap justify-between items-center gap-2"><span class="text-xs font-bold text-cyan-400">รายการอุปกรณ์ซ่อมบำรุง / เครื่องมือช่าง & อะไหล่ (เพิ่มได้หลายตัว + ถ่ายรูปรายชิ้น)</span><div class="flex gap-1.5"><button onclick="addQuickPresetEquipment()" class="px-2.5 py-1 rounded bg-indigo-600 hover:bg-indigo-500 text-white text-[11px] font-semibold">+ เพิ่มชุดเครื่องมือมาตรฐาน</button><button onclick="addBlankEquipItemRow()" class="px-2.5 py-1 rounded bg-emerald-600 hover:bg-emerald-500 text-white text-[11px] font-bold">+ เพิ่มรายการอุปกรณ์/อะไหล่</button></div></div>',
    '        <div class="overflow-x-auto"><table class="w-full text-left text-xs"><thead class="bg-slate-800 text-slate-300"><tr><th class="p-1.5">#</th><th class="p-1.5">รูป</th><th class="p-1.5">หมวดหมู่</th><th class="p-1.5">ชื่ออุปกรณ์ซ่อมบำรุง / อะไหล่ช่าง</th><th class="p-1.5 w-14 text-center">เบิก</th><th class="p-1.5 w-14 text-center">คืน</th><th class="p-1.5 w-14">หน่วย</th><th class="p-1.5">สถานะของอุปกรณ์</th><th class="p-1.5">หมายเหตุ</th><th class="p-1.5"></th></tr></thead><tbody id="eqItemsTableBody" class="divide-y divide-slate-800"></tbody></table></div>',
    '      </div>',
    '      <div class="flex flex-wrap justify-between items-center gap-2 pt-1"><input type="text" id="eqNote" placeholder="หมายเหตุใบเบิก-คืน..." class="flex-1 bg-slate-900 border border-slate-600 rounded-lg px-3 py-1.5 text-xs text-white"><button onclick="markCurrentEquipReturnedToday()" class="px-3 py-1.5 rounded-lg bg-cyan-500 hover:bg-cyan-400 text-slate-950 text-xs font-extrabold"><i class="fa-solid fa-rotate-left mr-1"></i>กดรับคืนครบ & บันทึก</button></div>',
    '    </div>',
    '    <div class="lg:col-span-5 bg-slate-800/90 border border-slate-700 rounded-2xl p-4 flex flex-col space-y-3">',
    '      <div class="flex justify-between items-center border-b border-slate-700 pb-2"><h3 class="text-sm font-bold text-amber-400"><i class="fa-solid fa-magnifying-glass-chart mr-1.5"></i>ตรวจสอบและดึงข้อมูลใบเบิก-คืนอุปกรณ์ขึ้นมาดู</h3><button onclick="loadEquipHistoryFromServer()" class="px-2.5 py-1 rounded bg-slate-700 hover:bg-slate-600 text-xs text-cyan-300"><i class="fa-solid fa-arrows-rotate mr-1"></i>รีเฟรช</button></div>',
    '      <div class="flex gap-2"><input type="text" id="eqSearchInput" oninput="renderEquipHistoryTable()" placeholder="ค้นหาผู้เบิก, ผู้นำมาคืน, ชื่ออุปกรณ์, วันที่..." class="flex-1 bg-slate-900 border border-slate-700 rounded-lg px-2.5 py-1.5 text-xs text-white"><select id="eqFilterStatus" onchange="renderEquipHistoryTable()" class="bg-slate-900 border border-slate-700 rounded-lg px-2 py-1.5 text-xs text-white"><option value="ALL">ทุกสถานะ</option><option value="กำลังเบิกขึ้นของเข้าทำงาน">กำลังขึ้นของ/ยังไม่คืน</option><option value="นำมาคืนครบแล้ว">นำมาคืนครบแล้ว</option><option value="คืนของเมื่อคืนเรียบร้อย">คืนของเมื่อคืนเรียบร้อย</option><option value="อุปกรณ์ชำรุด / ส่งซ่อม">ชำรุด/รอซ่อม</option></select></div>',
    '      <div class="overflow-y-auto max-h-[460px] border border-slate-700 rounded-xl flex-1"><table class="w-full text-left text-xs"><thead class="bg-slate-900 text-slate-300 sticky top-0"><tr><th class="p-2">เลขที่/วันที่เบิก-คืน</th><th class="p-2">ผู้เบิก / ผู้นำมาคืน</th><th class="p-2">สถานะ/รูป</th><th class="p-2 text-right">ดึงข้อมูล</th></tr></thead><tbody id="eqHistoryTableBody" class="divide-y divide-slate-800"></tbody></table></div>',
    '    </div>',
    '  </div>',
    '</section>',
    '<div id="a4ReportWorkspace" class="flex-1 flex flex-col lg:flex-row max-w-[1600px] w-full mx-auto overflow-hidden">',
    '  <aside class="no-print w-full lg:w-[430px] bg-slate-800/90 border-b lg:border-r border-slate-700 overflow-y-auto max-h-[58vh] lg:max-h-[calc(100vh-58px)] p-3 space-y-3 shrink-0">',
    '    <div class="bg-slate-900 rounded-xl p-3 border border-slate-700 space-y-2.5">',
    '      <div class="text-xs font-bold text-cyan-400">1. ข้อมูลเครื่องจักร & วันที่เข้าทำ</div>',
    '      <div class="grid grid-cols-2 gap-2">',
    '        <div><label class="block text-[11px] text-slate-300">รหัสเครื่องจักร *</label><input type="text" id="inputMachineCode" value="MCH-BLW-014" oninput="syncJobData()" class="w-full bg-slate-800 border border-slate-600 rounded px-2 py-1 text-xs text-white font-bold"></div>',
    '        <div><label class="block text-[11px] text-slate-300">วันที่เข้าทำ *</label><input type="date" id="inputServiceDate" oninput="syncJobData()" class="w-full bg-slate-800 border border-slate-600 rounded px-2 py-1 text-xs text-white"></div>',
    '        <div><label class="block text-[11px] text-slate-300">ชื่อเครื่องจักร/แผนก</label><input type="text" id="inputMachineName" value="ชุดมอเตอร์และโบลเวอร์ #2" oninput="syncJobData()" class="w-full bg-slate-800 border border-slate-600 rounded px-2 py-1 text-xs text-white"></div>',
    '        <div><label class="block text-[11px] text-slate-300">ช่างผู้ปฏิบัติงาน</label><input type="text" id="inputTechName" value="ทีมวิศวกรรมซ่อมบำรุง" oninput="syncJobData()" class="w-full bg-slate-800 border border-slate-600 rounded px-2 py-1 text-xs text-white"></div>',
    '      </div>',
    '    </div>',
    '    <div class="bg-slate-900 rounded-xl p-3 border border-slate-700 space-y-2.5">',
    '      <div class="flex justify-between items-center"><span class="text-xs font-bold text-amber-400">2. ตั้งค่ากระดาษ A4 ตารางรูป & ขนาดกรอบรูป</span><span id="currentPageIndicator" class="text-[11px] text-amber-300 font-bold">หน้า 1/1</span></div>',
    '      <div class="grid grid-cols-2 gap-2">',
    '        <button onclick="setOrientation(\'portrait\')" id="btnOrientPortrait" class="py-1 rounded text-xs font-semibold border border-cyan-500 bg-cyan-500/20 text-cyan-300">แนวตั้ง (Portrait)</button>',
    '        <button onclick="setOrientation(\'landscape\')" id="btnOrientLandscape" class="py-1 rounded text-xs font-semibold border border-slate-700 bg-slate-800 text-slate-300">แนวนอน (Landscape)</button>',
    '      </div>',
    '      <div><label class="block text-[11px] text-slate-300 mb-1">จำนวนรูปในหน้ากระดาษ A4 (4 / 6 / 8 / 12 รูป)</label>',
    '        <div class="grid grid-cols-4 gap-1.5">',
    '          <button onclick="setPageGridCount(4)" id="btnGrid4" class="py-1 rounded text-xs font-bold border border-cyan-500 bg-cyan-500/20 text-cyan-300">4 รูป</button>',
    '          <button onclick="setPageGridCount(6)" id="btnGrid6" class="py-1 rounded text-xs font-bold border border-slate-700 bg-slate-800 text-slate-300">6 รูป</button>',
    '          <button onclick="setPageGridCount(8)" id="btnGrid8" class="py-1 rounded text-xs font-bold border border-slate-700 bg-slate-800 text-slate-300">8 รูป</button>',
    '          <button onclick="setPageGridCount(12)" id="btnGrid12" class="py-1 rounded text-xs font-bold border border-slate-700 bg-slate-800 text-slate-300">12 รูป</button>',
    '        </div>',
    '      </div>',
    '      <div><label class="block text-[11px] text-slate-300 mb-1">ขนาดกรอบรูปภาพ (จัดกึ่งกลาง ไม่ล้นหน้า A4)</label>',
    '        <div class="grid grid-cols-3 sm:grid-cols-6 gap-1">',
    '          <button onclick="setFrameSize(\'4cm\')" data-size="4cm" class="size-btn py-1 rounded text-[11px] border border-slate-700 bg-slate-800 text-slate-300">4x4 ซม.</button>',
    '          <button onclick="setFrameSize(\'5cm\')" data-size="5cm" class="size-btn py-1 rounded text-[11px] border border-slate-700 bg-slate-800 text-slate-300">5x5 ซม.</button>',
    '          <button onclick="setFrameSize(\'6cm\')" data-size="6cm" class="size-btn py-1 rounded text-[11px] border border-cyan-500 bg-cyan-500/20 text-cyan-300">6x6 ซม.</button>',
    '          <button onclick="setFrameSize(\'7cm\')" data-size="7cm" class="size-btn py-1 rounded text-[11px] border border-slate-700 bg-slate-800 text-slate-300">7x7 ซม.</button>',
    '          <button onclick="setFrameSize(\'8cm\')" data-size="8cm" class="size-btn py-1 rounded text-[11px] border border-slate-700 bg-slate-800 text-slate-300">8x8 ซม.</button>',
    '          <button onclick="setFrameSize(\'9cm\')" data-size="9cm" class="size-btn py-1 rounded text-[11px] border border-slate-700 bg-slate-800 text-slate-300">9x9 ซม.</button>',
    '        </div>',
    '      </div>',
    '    </div>',
    '    <div class="bg-slate-900 rounded-xl p-3 border border-slate-700 space-y-2">',
    '      <div class="flex justify-between items-center"><span class="text-xs font-bold text-emerald-400">3. หัวข้อหน้ากระดาษ 1 - 7 (แยกสัดส่วนไม่ปนกัน)</span><button onclick="generateAll7SectionsPages()" class="text-[10px] bg-emerald-600/30 text-emerald-300 border border-emerald-500/40 px-2 py-0.5 rounded">+ สร้างครบ 7 หัวข้อ (7 หน้า)</button></div>',
    '      <select id="pageSectionSelect" onchange="changeActivePageSection(this.value)" class="w-full bg-slate-800 border border-slate-600 rounded px-2 py-1.5 text-xs text-white font-semibold">',
    '        <option value="SEC_1">หัวข้อที่ 1 : OVERHAUL MOTOR (ก่อนทำ - หลังทำ)</option>',
    '        <option value="SEC_2">หัวข้อที่ 2 : OVERHAUL BLOWER (ก่อนทำ - หลังทำ)</option>',
    '        <option value="SEC_3">หัวข้อที่ 3 : เปลี่ยนสายพาน และ ตั้ง Alignment Laser (ก่อนทำ - หลังทำ)</option>',
    '        <option value="SEC_4">หัวข้อที่ 4 : CLEANING (ก่อนทำ - หลังทำ)</option>',
    '        <option value="SEC_5">หัวข้อที่ 5 : PM (Preventive Maintenance)</option>',
    '        <option value="SEC_6">หัวข้อที่ 6 : วัดค่าความร้อน Temperature (°C) 3 ค่า & Vibration (mm/s) 3 ค่า</option>',
    '        <option value="SEC_7">หัวข้อที่ 7 : ข้อเสนอแนะ Comment / Alarm Code</option>',
    '      </select>',
    '      <div class="grid grid-cols-2 gap-1.5 pt-1">',
    '        <button onclick="openOrCreateSectionPage(\'SEC_6\')" class="py-1 px-2 rounded bg-rose-600/20 hover:bg-rose-600/40 border border-rose-500/40 text-rose-300 text-[11px] font-semibold">ไปที่หน้า หัวข้อที่ 6 (Temp/Vib)</button>',
    '        <button onclick="openOrCreateSectionPage(\'SEC_7\')" class="py-1 px-2 rounded bg-purple-600/20 hover:bg-purple-600/40 border border-purple-500/40 text-purple-300 text-[11px] font-semibold">ไปที่หน้า หัวข้อที่ 7 (Comment)</button>',
    '      </div>',
    '    </div>',
    '    <div class="bg-slate-900 rounded-xl p-3 border border-slate-700 space-y-2">',
    '      <div class="flex justify-between items-center"><span class="text-xs font-bold text-rose-400">4. ข้อมูลหัวข้อที่ 6 : ความร้อน (°C) 3 ค่า & ความสั่นสะเทือน (mm/s) 3 ค่า</span><button onclick="openOrCreateSectionPage(\'SEC_6\')" class="text-[10px] text-rose-300 underline">ดูหน้า 6</button></div>',
    '      <div class="grid grid-cols-3 gap-1.5">',
    '        <input type="text" id="tempVal1" value="48.5" oninput="syncJobData()" placeholder="Temp 1 °C" class="bg-slate-800 border border-slate-600 rounded px-2 py-1 text-xs text-white text-center">',
    '        <input type="text" id="tempVal2" value="46.2" oninput="syncJobData()" placeholder="Temp 2 °C" class="bg-slate-800 border border-slate-600 rounded px-2 py-1 text-xs text-white text-center">',
    '        <input type="text" id="tempVal3" value="44.8" oninput="syncJobData()" placeholder="Temp 3 °C" class="bg-slate-800 border border-slate-600 rounded px-2 py-1 text-xs text-white text-center">',
    '      </div>',
    '      <div class="grid grid-cols-3 gap-1.5">',
    '        <input type="text" id="vibVal1" value="1.85" oninput="syncJobData()" placeholder="Vib H mm/s" class="bg-slate-800 border border-slate-600 rounded px-2 py-1 text-xs text-white text-center">',
    '        <input type="text" id="vibVal2" value="1.42" oninput="syncJobData()" placeholder="Vib V mm/s" class="bg-slate-800 border border-slate-600 rounded px-2 py-1 text-xs text-white text-center">',
    '        <input type="text" id="vibVal3" value="1.15" oninput="syncJobData()" placeholder="Vib A mm/s" class="bg-slate-800 border border-slate-600 rounded px-2 py-1 text-xs text-white text-center">',
    '      </div>',
    '    </div>',
    '    <div class="bg-slate-900 rounded-xl p-3 border border-slate-700 space-y-2">',
    '      <div class="flex justify-between items-center"><span class="text-xs font-bold text-purple-400">5. ข้อมูลหัวข้อที่ 7 : ข้อเสนอแนะ Comment / Alarm Code</span><button onclick="openOrCreateSectionPage(\'SEC_7\')" class="text-[10px] text-purple-300 underline">ดูหน้า 7</button></div>',
    '      <input type="text" id="inputAlarmCode" value="ALM-00 (Normal)" oninput="syncJobData()" placeholder="Alarm Code" class="w-full bg-slate-800 border border-slate-600 rounded px-2 py-1 text-xs text-amber-300 font-bold">',
    '      <input type="text" id="inputComment" value="ดำเนินการ Overhaul ตั้งศูนย์เลเซอร์ และทำความสะอาดเรียบร้อย ค่าอยู่ในเกณฑ์มาตรฐาน" oninput="syncJobData()" placeholder="ข้อเสนอแนะ Comment" class="w-full bg-slate-800 border border-slate-600 rounded px-2 py-1 text-xs text-white">',
    '    </div>',
    '  </aside>',
    '  <main class="flex-1 flex flex-col bg-slate-950 overflow-hidden">',
    '    <div class="no-print bg-slate-900 border-b border-slate-800 px-3 py-2 flex flex-wrap items-center justify-between gap-2">',
    '      <div class="flex items-center gap-1.5 overflow-x-auto" id="pageTabsContainer"></div>',
    '      <div class="flex items-center gap-1.5">',
    '        <button onclick="addNewA4Page()" class="px-2.5 py-1 rounded bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-semibold">+ เพิ่มหน้า A4</button>',
    '        <button onclick="removeCurrentA4Page()" class="px-2.5 py-1 rounded bg-slate-800 hover:bg-rose-600 text-slate-300 text-xs">ลบหน้า</button>',
    '      </div>',
    '    </div>',
    '    <div id="a4PreviewViewport" class="flex-1 overflow-y-auto">',
    '      <div id="a4ScaleWrapper" class="a4-scale-wrapper"><div id="a4PagesContainer"></div></div>',
    '    </div>',
    '  </main>',
    '</div>',
    '<input type="file" id="singleSlotFileInput" accept="image/*" class="hidden" onchange="handleSingleSlotFile(event)">',
    '<input type="file" id="directCameraInput" accept="image/*" capture="environment" class="hidden" onchange="handleNativeCameraFallback(event)">',
    '<!-- Camera Modal (No Black Screen Guarantee) -->',
    '<div id="cameraModal" class="no-print fixed inset-0 z-50 hidden bg-slate-950/90 flex items-center justify-center p-3">',
    '  <div class="bg-slate-900 border border-slate-700 rounded-2xl max-w-md w-full p-4 space-y-3">',
    '    <div class="flex justify-between items-center">',
    '      <h3 class="text-sm font-bold text-white"><i class="fa-solid fa-camera text-indigo-400 mr-1.5"></i>ถ่ายรูปหน้างาน & สลับกล้อง</h3>',
    '      <button onclick="closeCameraModal()" class="w-7 h-7 rounded-lg bg-slate-800 text-slate-300 hover:bg-rose-600 hover:text-white"><i class="fa-solid fa-xmark"></i></button>',
    '    </div>',
    '    <div class="flex justify-between items-center text-xs bg-slate-800 px-3 py-1.5 rounded-lg border border-slate-700">',
    '      <span id="cameraTargetInfo" class="text-cyan-300 font-semibold">กำลังถ่ายลงช่องว่างถัดไป</span>',
    '      <span id="cameraFacingBadge" class="bg-indigo-500/20 text-indigo-300 px-2 py-0.5 rounded text-[11px] font-bold">กล้องหลัง (Back)</span>',
    '    </div>',
    '    <div class="relative bg-slate-800 rounded-xl overflow-hidden aspect-square max-h-[300px] mx-auto flex items-center justify-center border-2 border-slate-700">',
    '      <video id="cameraVideo" autoplay muted playsinline class="w-full h-full object-cover hidden"></video>',
    '      <div id="cameraFallbackBox" class="p-5 text-center space-y-2.5 flex flex-col items-center justify-center w-full h-full bg-gradient-to-b from-slate-800 to-slate-900">',
    '        <div class="w-16 h-16 rounded-full bg-indigo-600/20 border border-indigo-500/40 flex items-center justify-center text-indigo-400 text-2xl"><i class="fa-solid fa-camera-retro"></i></div>',
    '        <div class="text-xs font-bold text-white">พร้อมถ่ายรูปเข้าตารางกระดาษ A4</div>',
    '        <p class="text-[11px] text-slate-300">กดปุ่ม <b>"กดถ่ายรูป"</b> ด้านล่างเพื่อเปิดกล้องถ่ายรูป หรือกด <b>"สลับกล้อง"</b> เพื่อเลือกกล้องหน้า/หลัง</p>',
    '        <img id="lastCapturedPreview" class="hidden w-24 h-24 object-cover rounded-lg border-2 border-emerald-400 mx-auto">',
    '      </div>',
    '    </div>',
    '    <div class="grid grid-cols-3 gap-2 pt-1">',
    '      <button onclick="switchCameraFacing()" class="py-2.5 rounded-xl bg-slate-800 hover:bg-slate-700 text-white text-xs font-semibold border border-slate-600"><i class="fa-solid fa-camera-rotate text-cyan-400 mr-1"></i>สลับกล้อง</button>',
    '      <button onclick="capturePhotoFromVideo()" class="py-2.5 rounded-xl bg-indigo-600 hover:bg-indigo-500 text-white text-xs font-bold shadow-lg"><i class="fa-solid fa-circle-dot mr-1"></i>กดถ่ายรูป</button>',
    '      <button onclick="triggerDirectCameraInput()" class="py-2.5 rounded-xl bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-bold"><i class="fa-solid fa-mobile-screen mr-1"></i>เปิดกล้องมือถือ</button>',
    '    </div>',
    '  </div>',
    '</div>',
    '<!-- Slot Edit Modal -->',
    '<div id="slotEditModal" class="no-print fixed inset-0 z-50 hidden bg-slate-950/85 flex items-center justify-center p-3">',
    '  <div class="bg-slate-900 border border-slate-700 rounded-2xl max-w-md w-full p-4 space-y-3">',
    '    <div class="flex justify-between items-center"><h3 class="text-sm font-bold text-white">แก้ไขหัวข้อและหมายเหตุใต้รูปภาพ</h3><button onclick="closeSlotEditModal()" class="w-7 h-7 rounded bg-slate-800 text-white"><i class="fa-solid fa-xmark"></i></button></div>',
    '    <div class="grid grid-cols-2 gap-2">',
    '      <select id="modalSlotStatus" class="bg-slate-800 border border-slate-600 rounded px-2 py-1.5 text-xs text-white"><option value="ก่อนทำ (BEFORE)">ก่อนทำ (BEFORE)</option><option value="หลังทำ (AFTER)">หลังทำ (AFTER)</option><option value="PM / ตรวจสอบ">PM / ตรวจสอบ</option></select>',
    '      <input type="text" id="modalSlotTitle" class="bg-slate-800 border border-slate-600 rounded px-2 py-1.5 text-xs text-white" placeholder="หัวข้อย่อย">',
    '    </div>',
    '    <div><div class="flex justify-between mb-1"><span class="text-xs text-amber-300 font-semibold">หมายเหตุใต้รูป (เพิ่มได้หลายบรรทัด)</span><button onclick="addModalRemarkLine(\'\')" class="text-[11px] bg-emerald-600 text-white px-2 py-0.5 rounded">+ เพิ่มบรรทัด</button></div><div id="modalRemarksList" class="space-y-1.5 max-h-40 overflow-y-auto"></div></div>',
    '    <div class="flex justify-end gap-2"><button onclick="closeSlotEditModal()" class="px-3 py-1.5 rounded bg-slate-800 text-xs text-slate-300">ยกเลิก</button><button onclick="saveSlotEditModal()" class="px-4 py-1.5 rounded bg-cyan-500 text-slate-950 text-xs font-bold">บันทึก</button></div>',
    '  </div>',
    '</div>',
    '<!-- Google Sheets History Modal -->',
    '<div id="sheetsModal" class="no-print fixed inset-0 z-50 hidden bg-slate-950/90 flex items-center justify-center p-3">',
    '  <div class="bg-slate-900 border border-slate-700 rounded-2xl max-w-4xl w-full max-h-[88vh] flex flex-col p-4 space-y-3">',
    '    <div class="flex justify-between items-center border-b border-slate-800 pb-2">',
    '      <h3 class="text-sm font-bold text-emerald-400"><i class="fa-solid fa-table mr-1.5"></i>ประวัติใบงานใน Google Sheets</h3>',
    '      <div class="flex items-center gap-2"><a id="openDirectSheetLink" href="#" target="_blank" class="hidden px-2.5 py-1 rounded bg-emerald-600 text-white text-xs font-bold">เปิดตาราง Google Sheets จริง</a><button onclick="closeSheetsModal()" class="w-7 h-7 rounded bg-slate-800 text-white"><i class="fa-solid fa-xmark"></i></button></div>',
    '    </div>',
    '    <div class="overflow-auto flex-1 border border-slate-800 rounded-xl"><table class="w-full text-left text-xs"><thead class="bg-slate-800 text-slate-300"><tr><th class="p-2">วันที่</th><th class="p-2">รหัสเครื่องจักร</th><th class="p-2">สถานะ</th><th class="p-2">Temp 3 ค่า</th><th class="p-2">Vibration 3 ค่า</th><th class="p-2">Alarm / Comment</th><th class="p-2 text-right">ตรวจสอบ</th></tr></thead><tbody id="sheetsTableBody" class="divide-y divide-slate-800"></tbody></table></div>',
    '  </div>',
    '</div>',
    '<!-- Equipment Status Inspector Modal (กดดูสถานะจาก 4 การ์ด) -->',
    '<div id="eqStatusModal" class="no-print fixed inset-0 z-50 hidden bg-slate-950/90 flex items-center justify-center p-3">',
    '  <div class="bg-slate-900 border border-slate-700 rounded-2xl max-w-4xl w-full max-h-[88vh] flex flex-col p-4 space-y-3 shadow-2xl">',
    '    <div class="flex justify-between items-center border-b border-slate-800 pb-2">',
    '      <h3 id="eqStatusModalTitle" class="text-sm font-extrabold text-cyan-400">ตรวจสอบสถานะใบเบิก-คืนอุปกรณ์ซ่อมบำรุง</h3>',
    '      <button onclick="document.getElementById(\'eqStatusModal\').classList.add(\'hidden\')" class="w-7 h-7 rounded bg-slate-800 hover:bg-rose-600 text-white"><i class="fa-solid fa-xmark"></i></button>',
    '    </div>',
    '    <div class="grid grid-cols-2 sm:grid-cols-4 gap-2">',
    '      <button onclick="openEqStatusModal(\'ALL\')" class="py-1.5 px-2 rounded-lg bg-slate-800 hover:bg-cyan-500/20 border border-cyan-500/40 text-cyan-300 text-xs font-bold">ใบเบิก-จ่ายทั้งหมด</button>',
    '      <button onclick="openEqStatusModal(\'ACTIVE_OUT\')" class="py-1.5 px-2 rounded-lg bg-slate-800 hover:bg-amber-500/20 border border-amber-500/40 text-amber-300 text-xs font-bold">กำลังเบิกขึ้นของ (ยังไม่คืน)</button>',
    '      <button onclick="openEqStatusModal(\'RETURNED\')" class="py-1.5 px-2 rounded-lg bg-slate-800 hover:bg-emerald-500/20 border border-emerald-500/40 text-emerald-300 text-xs font-bold">นำของมาคืนครบแล้ว</button>',
    '      <button onclick="openEqStatusModal(\'REPAIR\')" class="py-1.5 px-2 rounded-lg bg-slate-800 hover:bg-rose-500/20 border border-rose-500/40 text-rose-300 text-xs font-bold">อุปกรณ์ซ่อมชำรุด / รอซ่อม</button>',
    '    </div>',
    '    <div id="eqStatusModalBody" class="overflow-y-auto flex-1 space-y-2.5 pr-1"></div>',
    '  </div>',
    '</div>',
    '<!-- Mobile Bottom Navigation Bar -->',
    '<nav class="no-print sm:hidden fixed bottom-0 inset-x-0 z-40 bg-slate-900/95 backdrop-blur border-t border-slate-700 px-2 py-1.5 grid grid-cols-3 gap-1 text-[10px] font-bold text-center shadow-2xl">',
    '  <button onclick="switchMainAppMode(\'A4_REPORT\')" class="py-1.5 rounded-lg bg-cyan-500/20 text-cyan-300 flex flex-col items-center justify-center gap-0.5"><i class="fa-solid fa-file-invoice text-sm"></i><span>รายงาน A4</span></button>',
    '  <button onclick="triggerBatchUpload()" class="py-1.5 rounded-lg bg-blue-500/20 text-blue-300 flex flex-col items-center justify-center gap-0.5"><i class="fa-solid fa-images text-sm"></i><span>เพิ่มหลายรูป</span></button>',
    '  <button onclick="switchMainAppMode(\'EQUIPMENT\')" class="py-1.5 rounded-lg bg-emerald-500/20 text-emerald-300 flex flex-col items-center justify-center gap-0.5"><i class="fa-solid fa-toolbox text-sm"></i><span>เบิก-คืนของ</span></button>',
    '</nav>',
    '<div id="toastBox" class="no-print fixed bottom-4 right-4 z-[10000] hidden bg-slate-900 border border-cyan-500 text-white px-4 py-2.5 rounded-xl shadow-2xl text-xs font-semibold"><span id="toastMessage"></span></div>',
    '<script>',
    'const SECTION_DEFINITIONS = {',
    '  SEC_1: { code:"SEC_1", num:"1", title:"1. OVERHAUL MOTOR (ก่อนทำ - หลังทำ)", badgeColor:"#1d4ed8", bgLight:"#eff6ff", borderColor:"#93c5fd", defaultSub:"Overhaul Motor" },',
    '  SEC_2: { code:"SEC_2", num:"2", title:"2. OVERHAUL BLOWER (ก่อนทำ - หลังทำ)", badgeColor:"#0369a1", bgLight:"#f0f9ff", borderColor:"#7dd3fc", defaultSub:"Overhaul Blower" },',
    '  SEC_3: { code:"SEC_3", num:"3", title:"3. เปลี่ยนสายพาน และ ตั้ง Alignment Laser (ก่อนทำ - หลังทำ)", badgeColor:"#b45309", bgLight:"#fffbeb", borderColor:"#fcd34d", defaultSub:"เปลี่ยนสายพาน / Alignment Laser" },',
    '  SEC_4: { code:"SEC_4", num:"4", title:"4. CLEANING (ก่อนทำ - หลังทำ)", badgeColor:"#047857", bgLight:"#ecfdf5", borderColor:"#6ee7b7", defaultSub:"Cleaning" },',
    '  SEC_5: { code:"SEC_5", num:"5", title:"5. PM (Preventive Maintenance)", badgeColor:"#6d28d9", bgLight:"#f5f3ff", borderColor:"#c4b5fd", defaultSub:"ตรวจเช็ค PM" },',
    '  SEC_6: { code:"SEC_6", num:"6", title:"6. วัดค่าความร้อน Temperature (°C) 3 ค่า และ ความสั่นสะเทือน Vibration (mm/s) 3 ค่า", badgeColor:"#be123c", bgLight:"#fff1f2", borderColor:"#fda4af", defaultSub:"วัดค่าความร้อน / ความสั่นสะเทือน" },',
    '  SEC_7: { code:"SEC_7", num:"7", title:"7. ข้อเสนอแนะ Comment / Alarm Code", badgeColor:"#7e22ce", bgLight:"#faf5ff", borderColor:"#d8b4fe", defaultSub:"ข้อเสนอแนะ / Alarm Code" }',
    '};',
    'let appState = { jobStatus:"OPEN", machineCode:"MCH-BLW-014", serviceDate:new Date().toISOString().split("T")[0], machineName:"ชุดมอเตอร์และโบลเวอร์ #2", techName:"ทีมวิศวกรรมซ่อมบำรุง", orientation:"portrait", frameSize:"6cm", activePageIndex:0, tempValues:["48.5","46.2","44.8"], vibValues:["1.85","1.42","1.15"], alarmCode:"ALM-00 (Normal)", comment:"ดำเนินการ Overhaul ตั้งศูนย์เลเซอร์ และทำความสะอาดเรียบร้อย ค่าอยู่ในเกณฑ์มาตรฐาน", pages:[], savedSheetsRecords:[] };',
    'let targetSlotRef = null; let cameraStream = null; let isLiveCameraReady = false; let currentCameraFacing = "environment";',
    'function createPageObject(secCode, gridCount) {',
    '  const sec = SECTION_DEFINITIONS[secCode] || SECTION_DEFINITIONS.SEC_1; const slots = [];',
    '  for (let i=0; i<gridCount; i++) {',
    '    const isBefore = (i%2===0);',
    '    const defaultStatus = (secCode==="SEC_5" || secCode==="SEC_6" || secCode==="SEC_7") ? "PM / ตรวจสอบ" : (isBefore ? "ก่อนทำ (BEFORE)" : "หลังทำ (AFTER)");',
    '    slots.push({ sectionTitle: sec.title.split(" (")[0], statusTag: defaultStatus, subTitle: sec.defaultSub + " #" + (i+1), remarks:["ตรวจสอบเรียบร้อย"], imageSrc:null });',
    '  }',
    '  return { sectionCode:secCode, gridCount:gridCount, slots:slots };',
    '}',
    'window.onload = function() {',
    '  document.getElementById("inputServiceDate").value = appState.serviceDate;',
    '  appState.pages = [createPageObject("SEC_1", 4)];',
    '  setupBeautifulPwaIconAndManifest();',
    '  initEquipmentWorkspace();',
    '  renderAll(); loadHistoryFromServer(); loadEquipHistoryFromServer();',
    '};',
    'let equipState = { checkoutPhoto: null, returnPhoto: null, currentItems: [], records: [] };',
    'let cameraModeContext = "REPORT"; let cameraEquipItemIdx = 0; let deferredPwaPrompt = null; let generatedAppIconUrl = ""; let serverWebAppUrl = ""; let currentSelectedIconTheme = 1;',
    'const CDN_3D_ICONS = {',
    '  1: "https://cdn-icons-png.flaticon.com/512/10484/10484536.png",',
    '  2: "https://cdn-icons-png.flaticon.com/512/8654/8654197.png",',
    '  3: "https://cdn-icons-png.flaticon.com/512/9912/9912615.png",',
    '  4: "https://cdn-icons-png.flaticon.com/512/10691/10691802.png"',
    '};',
    'window.addEventListener("beforeinstallprompt", function(e){ e.preventDefault(); deferredPwaPrompt = e; });',
    'function createBeautifulAppIconDataUrl(size, themeId) {',
    '  size = size || 512; themeId = themeId || currentSelectedIconTheme || 1;',
    '  const themes = {',
    '    1: { bg1:"#090d16", bg2:"#1e3a8a", bg3:"#0284c7", ring:"#38bdf8", gear:"#fbbf24", tooth:"#f59e0b", badge:"#10b981", center:"A4", label:"MAINT PRO" },',
    '    2: { bg1:"#064e3b", bg2:"#047857", bg3:"#10b981", ring:"#6ee7b7", gear:"#fde047", tooth:"#eab308", badge:"#38bdf8", center:"EQ", label:"EQUIP PRO" },',
    '    3: { bg1:"#2e1065", bg2:"#581c87", bg3:"#7e22ce", ring:"#e879f9", gear:"#38bdf8", tooth:"#0284c7", badge:"#fbbf24", center:"PM", label:"MOTOR LASER" },',
    '    4: { bg1:"#450a0a", bg2:"#991b1b", bg3:"#dc2626", ring:"#fca5a5", gear:"#fbbf24", tooth:"#d97706", badge:"#facc15", center:"PRO", label:"ENGINEER A4" }',
    '  };',
    '  const t = themes[themeId] || themes[1];',
    '  const c = document.createElement("canvas"); c.width = size; c.height = size; const ctx = c.getContext("2d");',
    '  const r = size * 0.22;',
    '  const bgGrad = ctx.createLinearGradient(0, 0, size, size);',
    '  bgGrad.addColorStop(0, t.bg1); bgGrad.addColorStop(0.5, t.bg2); bgGrad.addColorStop(1, t.bg3);',
    '  ctx.beginPath(); ctx.moveTo(r, 0); ctx.lineTo(size-r, 0); ctx.quadraticCurveTo(size, 0, size, r); ctx.lineTo(size, size-r); ctx.quadraticCurveTo(size, size, size-r, size); ctx.lineTo(r, size); ctx.quadraticCurveTo(0, size, 0, size-r); ctx.lineTo(0, r); ctx.quadraticCurveTo(0, 0, r, 0); ctx.closePath();',
    '  ctx.fillStyle = bgGrad; ctx.fill();',
    '  ctx.lineWidth = size * 0.03; ctx.strokeStyle = t.ring; ctx.stroke();',
    '  const cx = size * 0.5, cy = size * 0.41, outerR = size * 0.24, innerR = size * 0.15;',
    '  ctx.save(); ctx.translate(cx, cy);',
    '  for (let i = 0; i < 8; i++) {',
    '    ctx.rotate(Math.PI / 4); ctx.fillStyle = t.tooth;',
    '    ctx.fillRect(-size*0.04, -outerR - size*0.042, size*0.08, size*0.072);',
    '  }',
    '  ctx.beginPath(); ctx.arc(0, 0, outerR, 0, Math.PI * 2); ctx.fillStyle = t.gear; ctx.fill();',
    '  ctx.lineWidth = size * 0.02; ctx.strokeStyle = "#ffffff"; ctx.stroke();',
    '  ctx.beginPath(); ctx.arc(0, 0, innerR, 0, Math.PI * 2); ctx.fillStyle = t.bg1; ctx.fill();',
    '  ctx.fillStyle = t.ring; ctx.font = "900 " + Math.round(size * 0.12) + "px sans-serif"; ctx.textAlign = "center"; ctx.textBaseline = "middle"; ctx.fillText(t.center, 0, 2);',
    '  ctx.restore();',
    '  ctx.fillStyle = t.badge; ctx.beginPath(); ctx.roundRect ? ctx.roundRect(size*0.11, size*0.72, size*0.78, size*0.175, size*0.05) : ctx.rect(size*0.11, size*0.72, size*0.78, size*0.175); ctx.fill();',
    '  ctx.fillStyle = "#0f172a"; ctx.font = "900 " + Math.round(size * 0.082) + "px sans-serif"; ctx.textAlign = "center"; ctx.textBaseline = "middle";',
    '  ctx.fillText(t.label, size * 0.5, size * 0.81);',
    '  return c.toDataURL("image/png");',
    '}',
    'function updateAllAppIconPreviews(iconUrl) {',
    '  ["headerBrandAppIcon"].forEach(function(id){',
    '    const el = document.getElementById(id); if (el) el.src = iconUrl;',
    '  });',
    '}',
    'function cycleHeaderAppIcon() {',
    '  currentSelectedIconTheme = (currentSelectedIconTheme % 4) + 1;',
    '  generatedAppIconUrl = createBeautifulAppIconDataUrl(512, currentSelectedIconTheme);',
    '  updateAllAppIconPreviews(generatedAppIconUrl);',
    '  renderA4Pages();',
    '  showToast("เปลี่ยนสไตล์ไอคอนแอป 3 มิติเรียบร้อยแล้ว!");',
    '}',
    'function setupBeautifulPwaIconAndManifest() {',
    '  generatedAppIconUrl = createBeautifulAppIconDataUrl(512, 1);',
    '  updateAllAppIconPreviews(generatedAppIconUrl);',
    '  [1,2,3,4].forEach(function(tid){ const el = document.getElementById("iconThemePreview" + tid); if (el) el.src = createBeautifulAppIconDataUrl(160, tid); });',
    '  const fav = document.createElement("link"); fav.id = "dynamicFavLink"; fav.rel = "icon"; fav.type = "image/png"; fav.href = generatedAppIconUrl; document.head.appendChild(fav);',
    '  const apple = document.createElement("link"); apple.id = "dynamicAppleLink"; apple.rel = "apple-touch-icon"; apple.href = generatedAppIconUrl; document.head.appendChild(apple);',
    '  const sameOriginUrl = window.location.href;',
    '  const manifest = { id:window.location.pathname || "/", name:"MAINT & EQUIP PRO (ระบบเบิก-คืนอุปกรณ์ & รายงาน A4)", short_name:"MAINT PRO", start_url:sameOriginUrl, scope:window.location.origin + "/", display:"standalone", background_color:"#0f172a", theme_color:"#0f172a", icons:[{ src:generatedAppIconUrl, sizes:"192x192", type:"image/png", purpose:"any" }, { src:generatedAppIconUrl, sizes:"512x512", type:"image/png", purpose:"any" }, { src:generatedAppIconUrl, sizes:"512x512", type:"image/png", purpose:"maskable" }] };',
    '  const mLink = document.createElement("link"); mLink.rel = "manifest"; mLink.href = "data:application/manifest+json;charset=utf-8," + encodeURIComponent(JSON.stringify(manifest)); document.head.appendChild(mLink);',
    '  registerInlineServiceWorker();',
    '  if (typeof google !== "undefined" && google.script && google.script.run) {',
    '    google.script.run.saveCustomAppIconToDrive(generatedAppIconUrl, CDN_3D_ICONS[1]);',
    '  }',
    '}',
    'function selectCustomIconTheme(themeId) {',
    '  currentSelectedIconTheme = themeId;',
    '  generatedAppIconUrl = createBeautifulAppIconDataUrl(512, themeId);',
    '  updateAllAppIconPreviews(generatedAppIconUrl);',
    '  showToast("เลือกไอคอน 3 มิติแบบที่ " + themeId + " แล้ว กดปุ่มสีส้มเพื่อบันทึกเป็นไอคอนหลักได้เลย!");',
    '}',
    'function handleCustomIconUpload(e) {',
    '  const f = e.target.files[0]; if (!f) return;',
    '  const reader = new FileReader();',
    '  reader.onload = function(ev) {',
    '    const img = new Image();',
    '    img.onload = function() {',
    '      const c = document.createElement("canvas"); c.width = 512; c.height = 512;',
    '      const ctx = c.getContext("2d"); ctx.drawImage(img, 0, 0, 512, 512);',
    '      generatedAppIconUrl = c.toDataURL("image/png");',
    '      updateAllAppIconPreviews(generatedAppIconUrl);',
    '      showToast("อัปโหลดรูปไอคอนของคุณสำเร็จ! กดปุ่มสีส้มเพื่อบันทึกเป็นไอคอนแอปได้เลย");',
    '    };',
    '    img.src = ev.target.result;',
    '  };',
    '  reader.readAsDataURL(f);',
    '}',
    'function applySelectedIconToServer() {',
    '  const btn = document.getElementById("btnSaveServerIcon");',
    '  if (btn) btn.innerHTML = "<i class=\\"fa-solid fa-spinner fa-spin\\"></i> กำลังบันทึกไอคอน 3 มิติลงระบบ...";',
    '  if (typeof google !== "undefined" && google.script && google.script.run) {',
    '    google.script.run.withSuccessHandler(function(res){',
    '      if (btn) btn.innerHTML = "<i class=\\"fa-solid fa-circle-check\\"></i> บันทึกไอคอน 3 มิติลงระบบสำเร็จ!";',
    '      const target = (res && res.webAppUrl) ? res.webAppUrl : getBestInstallUrl();',
    '      const reloadLink = document.getElementById("reloadWithNewIconLink");',
    '      if (reloadLink) { reloadLink.href = target; reloadLink.classList.remove("hidden"); }',
    '      showToast("✅ เปลี่ยนไอคอนแอปหลักเรียบร้อยแล้ว! กดปุ่มสีเขียวเพื่อเปิดแอปด้วยไอคอนใหม่ได้เลย");',
    '    }).saveCustomAppIconToDrive(generatedAppIconUrl, CDN_3D_ICONS[currentSelectedIconTheme]);',
    '  } else {',
    '    if (btn) btn.innerHTML = "<i class=\\"fa-solid fa-circle-check\\"></i> อัปเดตไอคอน 3 มิติเรียบร้อยแล้ว!";',
    '    showToast("✅ อัปเดตไอคอน 3 มิติเรียบร้อยแล้ว!");',
    '  }',
    '}',
    'async function buildAndDownloadAndroidApkPackage() {',
    '  if (deferredPwaPrompt) {',
    '    try { deferredPwaPrompt.prompt(); await deferredPwaPrompt.userChoice; deferredPwaPrompt = null; markAppAsInstalledUI(); return; } catch(e) {}',
    '  }',
    '  copyWebAppInstallLink();',
    '  downloadUnlockPwaSingleFile();',
    '  const detailBox = document.getElementById("apkBuilderDetailBox");',
    '  if (detailBox) detailBox.classList.remove("hidden");',
    '  showToast("ดาวน์โหลดชุดติดตั้ง WebAPK และคัดลอกลิงก์สำหรับคอมไพล์ .APK เรียบร้อยแล้ว!");',
    '}',
    'function installIosIpaWebClipFromServer() {',
    '  const targetUrl = getBestInstallUrl();',
    '  const btn = document.getElementById("btnIosIpaInstall");',
    '  if (typeof google !== "undefined" && google.script && google.script.run) {',
    '    if (btn) btn.innerHTML = "<div class=\\"text-xs font-extrabold\\"><i class=\\"fa-solid fa-spinner fa-spin\\"></i> กำลังสร้างโปรไฟล์ iOS...</div>";',
    '    google.script.run.withSuccessHandler(function(res){',
    '      if (btn) btn.innerHTML = "<div class=\\"text-xs font-extrabold\\"><i class=\\"fa-brands fa-apple\\"></i> ติดตั้งโปรไฟล์ .IPA / iOS</div><div class=\\"text-[10px] text-sky-100\\">พร้อมติดตั้งลง iPhone / iPad</div>";',
    '      if (res && res.status === "SUCCESS" && res.downloadUrl) {',
    '        window.top.location.href = res.downloadUrl;',
    '        showToast("กำลังดาวน์โหลดโปรไฟล์แอป iOS ไปที่ การตั้งค่า (Settings) > ดาวน์โหลดโปรไฟล์แล้ว เพื่อกดติดตั้ง!");',
    '      } else {',
    '        installIosWebClipProfile();',
    '      }',
    '    }).createIosInstallerFileInDrive(targetUrl, generatedAppIconUrl);',
    '  } else {',
    '    installIosWebClipProfile();',
    '  }',
    '}',
    'function downloadUnlockPwaSingleFile() {',
    '  const targetUrl = getBestInstallUrl();',
    '  const pwaHtml = "<!DOCTYPE html><html lang=\\"th\\"><head><meta charset=\\"UTF-8\\"><meta name=\\"viewport\\" content=\\"width=device-width,initial-scale=1,maximum-scale=5\\"><meta name=\\"theme-color\\" content=\\"#0f172a\\"><meta name=\\"apple-mobile-web-app-capable\\" content=\\"yes\\"><meta name=\\"mobile-web-app-capable\\" content=\\"yes\\"><title>MAINT & EQUIP PRO</title><link rel=\\"icon\\" type=\\"image/png\\" href=\\"" + generatedAppIconUrl + "\\"><link rel=\\"apple-touch-icon\\" href=\\"" + generatedAppIconUrl + "\\"><script>const m={name:\\"MAINT & EQUIP PRO\\",short_name:\\"MAINT PRO\\",start_url:window.location.href,display:\\"standalone\\",background_color:\\"#0f172a\\",theme_color:\\"#0f172a\\",icons:[{src:\\"" + generatedAppIconUrl + "\\",sizes:\\"192x192\\",type:\\"image/png\\",purpose:\\"any\\"},{src:\\"" + generatedAppIconUrl + "\\",sizes:\\"512x512\\",type:\\"image/png\\",purpose:\\"any\\"}]};const l=document.createElement(\\"link\\");l.rel=\\"manifest\\";l.href=\\"data:application/manifest+json;charset=utf-8,\\"+encodeURIComponent(JSON.stringify(m));document.head.appendChild(l);<\\/script><style>html,body,iframe{margin:0;padding:0;width:100%;height:100%;border:none;background:#0f172a;overflow:hidden;}</style></head><body><iframe src=\\"" + targetUrl + "\\" allow=\\"camera;microphone;fullscreen;clipboard-write\\"></iframe></body></html>";',
    '  const dataUri = "data:text/html;charset=utf-8," + encodeURIComponent(pwaHtml);',
    '  triggerSafeDownload(dataUri, "MAINT_PRO_WebAPK_Installer.html");',
    '}',
    'function registerInlineServiceWorker() {',
    '  if ("serviceWorker" in navigator) {',
    '    try {',
    '      const swCode = "self.addEventListener(\\"install\\",function(e){self.skipWaiting();});self.addEventListener(\\"activate\\",function(e){e.waitUntil(self.clients.claim());});self.addEventListener(\\"fetch\\",function(e){});";',
    '      const blob = new Blob([swCode], { type:"application/javascript" });',
    '      navigator.serviceWorker.register(URL.createObjectURL(blob)).catch(function(){});',
    '    } catch(e) {}',
    '  }',
    '}',
    'function getBestInstallUrl() {',
    '  if (serverWebAppUrl && serverWebAppUrl.indexOf("http") === 0) return serverWebAppUrl;',
    '  const inp = document.getElementById("installWebAppUrlInput"); if (inp && inp.value.trim().indexOf("http") === 0) return inp.value.trim();',
    '  return window.location.href;',
    '}',
    'function openInstallAppModal() {',
    '  autoInstallAndCreateShortcut();',
    '}',
    'async function autoInstallAndCreateShortcut() {',
    '  const modal = document.getElementById("installAppModal");',
    '  if (modal) modal.classList.remove("hidden");',
    '  const bestUrl = getBestInstallUrl();',
    '  const inp = document.getElementById("installWebAppUrlInput"); if (inp && !inp.value) inp.value = bestUrl;',
    '  updateQrCodeFromInput(bestUrl);',
    '  updateDirectInstallLinks(bestUrl);',
    '  detectUserDeviceForInstallUI();',
    '  if (deferredPwaPrompt) {',
    '    try {',
    '      deferredPwaPrompt.prompt();',
    '      const choice = await deferredPwaPrompt.userChoice;',
    '      deferredPwaPrompt = null;',
    '      if (choice && choice.outcome === "accepted") {',
    '        markAppAsInstalledUI();',
    '        showToast("✅ ติดตั้งไอคอนแอปลงหน้าจอโฮมของโทรศัพท์สำเร็จ!");',
    '        return;',
    '      }',
    '    } catch(e) {}',
    '  }',
    '  startInstantAppInstall(false);',
    '}',
    'function detectUserDeviceForInstallUI() {',
    '  const sub = document.getElementById("detectedOsSubtitle");',
    '  const ua = navigator.userAgent || "";',
    '  if (/Android/i.test(ua)) {',
    '    if (sub) sub.textContent = "ตรวจพบมือถือ Android: กด ⋮ มุมขวาบน > เพิ่มลงในหน้าจอหลัก > แตะ สร้างทางลัด";',
    '  } else if (/iPhone|iPad|iPod/i.test(ua)) {',
    '    if (sub) sub.textContent = "ตรวจพบ iPhone/iPad: กดปุ่มติดตั้งโปรไฟล์ .IPA/iOS หรือกดแชร์ > เพิ่มไปยังหน้าจอโฮม";',
    '  } else {',
    '    if (sub) sub.textContent = "ตรวจพบคอมพิวเตอร์: สแกน QR Code ด้านล่างด้วยมือถือ หรือกดดาวน์โหลดตัวเปิดแอป";',
    '  }',
    '}',
    'function updateDirectInstallLinks(urlStr) {',
    '  const clean = (urlStr || window.location.href).trim();',
    '  const andLink = document.getElementById("androidChromeDirectLink");',
    '  if (andLink) andLink.href = clean;',
    '  const iosLink = document.getElementById("iosDirectProfileDownloadLink");',
    '  if (iosLink) iosLink.href = buildIosBase64ProfileUri(clean);',
    '}',
    'function detectUserDeviceForInstallUI() {',
    '  const sub = document.getElementById("detectedOsSubtitle");',
    '  const ua = navigator.userAgent || "";',
    '  if (/Android/i.test(ua)) {',
    '    if (sub) sub.textContent = "ตรวจพบมือถือ Android: กด ⋮ มุมขวาบน > เพิ่มลงในหน้าจอหลัก > แตะ สร้างทางลัด";',
    '  } else if (/iPhone|iPad|iPod/i.test(ua)) {',
    '    if (sub) sub.textContent = "ตรวจพบ iPhone/iPad: กดปุ่มติดตั้งโปรไฟล์ .IPA/iOS หรือกดแชร์ > เพิ่มไปยังหน้าจอโฮม";',
    '  } else {',
    '    if (sub) sub.textContent = "ตรวจพบคอมพิวเตอร์: สแกน QR Code ด้านล่างด้วยมือถือ หรือกดสร้างไฟล์ .APK / .IPA";',
    '  }',
    '}',
    'function updateDirectInstallLinks(urlStr) {',
    '  const clean = (urlStr || window.location.href).trim();',
    '  const andLink = document.getElementById("androidChromeDirectLink");',
    '  if (andLink) {',
    '    if (/Android/i.test(navigator.userAgent) && clean.indexOf("http") === 0) {',
    '      const noProto = clean.replace(/^https?:\\/\\//i, "");',
    '      andLink.href = "intent://" + noProto + "#Intent;scheme=https;package=com.android.chrome;end";',
    '    } else {',
    '      andLink.href = clean;',
    '    }',
    '  }',
    '}',
    'function buildIosBase64ProfileUri(targetUrl) {',
    '  const base64Icon = generatedAppIconUrl.replace(/^data:image\\/png;base64,/, "");',
    '  const uuid1 = "A4MAINT-1111-2222-3333-" + Date.now();',
    '  const uuid2 = "A4MAINT-4444-5555-6666-" + Date.now();',
    '  const xml = "<?xml version=\\"1.0\\" encoding=\\"UTF-8\\"?><!DOCTYPE plist PUBLIC \\"-//Apple//DTD PLIST 1.0//EN\\" \\"http://www.apple.com/DTDs/PropertyList-1.0.dtd\\"><plist version=\\"1.0\\"><dict><key>PayloadContent</key><array><dict><key>FullScreen</key><true/><key>Icon</key><data>" + base64Icon + "</data><key>IsRemovable</key><true/><key>Label</key><string>MAINT PRO</string><key>PayloadDescription</key><string>Install MAINT PRO Home Screen Icon</string><key>PayloadDisplayName</key><string>MAINT PRO</string><key>PayloadIdentifier</key><string>com.maintpro.webclip." + uuid1 + "</string><key>PayloadType</key><string>com.apple.webClip.managed</string><key>PayloadUUID</key><string>" + uuid1 + "</string><key>PayloadVersion</key><integer>1</integer><key>Precomposed</key><true/><key>URL</key><string>" + targetUrl + "</string></dict></array><key>PayloadDisplayName</key><string>MAINT PRO Home Screen Icon</string><key>PayloadIdentifier</key><string>com.maintpro.profile." + uuid2 + "</string><key>PayloadRemovalDisallowed</key><false/><key>PayloadType</key><string>Configuration</string><key>PayloadUUID</key><string>" + uuid2 + "</string><key>PayloadVersion</key><integer>1</integer></dict></plist>";',
    '  const b64 = btoa(unescape(encodeURIComponent(xml)));',
    '  return "data:application/x-apple-aspen-config;base64," + b64;',
    '}',
    'async function handleAndroidDirectClick(e) {',
    '  if (deferredPwaPrompt) {',
    '    e.preventDefault();',
    '    deferredPwaPrompt.prompt();',
    '    await deferredPwaPrompt.userChoice;',
    '    deferredPwaPrompt = null;',
    '    markAppAsInstalledUI();',
    '    return false;',
    '  }',
    '  showToast("แตะปุ่ม ⋮ มุมขวาบนของจอ > กด เพิ่มลงในหน้าจอหลัก > แตะบรรทัดที่ 2 (สร้างทางลัด) ไอคอนจะขึ้นหน้าจอมือถือทันที!");',
    '  return true;',
    '}',
    'function markAppAsInstalledUI() {',
    '  const btnTxt = document.getElementById("headerInstallBtnText");',
    '  const btn = document.getElementById("headerInstallAppBtn");',
    '  if (btnTxt) btnTxt.textContent = "✅ ติดตั้งทางลัดแอปแล้ว";',
    '  if (btn) btn.classList.remove("animate-pulse");',
    '  const pill = document.getElementById("installAppStatusPill");',
    '  if (pill) { pill.textContent = "✅ ติดตั้งลงหน้าจอโฮมสำเร็จแล้ว"; pill.className = "inline-block px-2.5 py-0.5 rounded-full text-[10px] font-extrabold bg-cyan-400 text-slate-950"; }',
    '}',
    'function closeInstallAppModal() {',
    '  const modal = document.getElementById("installAppModal");',
    '  if (modal) modal.classList.add("hidden");',
    '}',
    'function updateQrCodeFromInput(urlStr) {',
    '  const qr = document.getElementById("qrCodeInstallImg");',
    '  const clean = (urlStr || window.location.href).trim();',
    '  if (qr) qr.src = "https://api.qrserver.com/v1/create-qr-code/?size=240x240&data=" + encodeURIComponent(clean);',
    '  updateDirectInstallLinks(clean);',
    '}',
    'function copyWebAppInstallLink() {',
    '  const inp = document.getElementById("installWebAppUrlInput"); if (!inp) return;',
    '  inp.select(); document.execCommand("copy"); showToast("คัดลอกลิงก์เว็บแอปเรียบร้อย ส่งเข้า LINE เพื่อเปิดในมือถือได้เลย!");',
    '}',
    'function downloadHighResAppIcon() {',
    '  triggerSafeDownload(generatedAppIconUrl, "MaintEquipPro_Icon_512.png");',
    '  showToast("ดาวน์โหลดรูปไอคอนแอป 512x512 PNG สำเร็จ!");',
    '}',
    'function triggerSafeDownload(dataOrBlobUrl, fileName) {',
    '  try {',
    '    const a = document.createElement("a");',
    '    a.href = dataOrBlobUrl;',
    '    a.download = fileName;',
    '    a.target = "_blank";',
    '    document.body.appendChild(a);',
    '    a.click();',
    '    setTimeout(function(){ document.body.removeChild(a); }, 300);',
    '  } catch (e) {',
    '    window.open(dataOrBlobUrl, "_blank");',
    '  }',
    '}',
    'let installTimerRef = null;',
    'async function startInstantAppInstall(autoDownloadShortcut) {',
    '  if (deferredPwaPrompt) {',
    '    try { deferredPwaPrompt.prompt(); await deferredPwaPrompt.userChoice; deferredPwaPrompt = null; markAppAsInstalledUI(); showToast("ติดตั้งแอปลงหน้าจอโทรศัพท์สำเร็จ!"); return; } catch(e) {}',
    '  }',
    '  const progBox = document.getElementById("installProgressBox");',
    '  const progBar = document.getElementById("installProgressBar");',
    '  const progPct = document.getElementById("installProgressPct");',
    '  const progTxt = document.getElementById("installProgressText");',
    '  const doneHint = document.getElementById("installDoneHint");',
    '  if (progBox) progBox.classList.remove("hidden");',
    '  if (installTimerRef) clearInterval(installTimerRef);',
    '  let pct = 0;',
    '  installTimerRef = setInterval(function(){',
    '    pct += 25;',
    '    if (pct > 100) pct = 100;',
    '    if (progBar) progBar.style.width = pct + "%";',
    '    if (progPct) progPct.textContent = pct + "%";',
    '    if (pct === 25 && progTxt) progTxt.textContent = "กำลังสร้างไอคอน 3D MAINT PRO (512x512)...";',
    '    if (pct === 50 && progTxt) progTxt.textContent = "กำลังผูกทางลัดเข้ากับหน้าจออุปกรณ์...";',
    '    if (pct >= 100) {',
    '      clearInterval(installTimerRef);',
    '      markAppAsInstalledUI();',
    '      if (progTxt) progTxt.textContent = "✅ ติดตั้งแอป & สร้างไฟล์ทางลัดสำเร็จแล้ว!";',
    '      if (doneHint) { doneHint.innerHTML = "✅ <b>สร้างทางลัด MaintEquipPro_App.html สำเร็จ!</b> แตะเปิดไฟล์ที่ดาวน์โหลดเพื่อเข้าแอปได้ทันที หรือกดปุ่ม <b>Android / iPhone</b> ด้านล่างเพื่อปักหมุดลงหน้าจอโฮม"; }',
    '      if (autoDownloadShortcut) downloadMobileAppLauncherFile();',
    '    }',
    '  }, 150);',
    '}',
    'function downloadMobileAppLauncherFile() {',
    '  const targetUrl = getBestInstallUrl();',
    '  const html = "<!DOCTYPE html><html lang=\\"th\\"><head><meta charset=\\"UTF-8\\"><meta name=\\"viewport\\" content=\\"width=device-width,initial-scale=1,maximum-scale=1,user-scalable=no\\"><meta name=\\"apple-mobile-web-app-capable\\" content=\\"yes\\"><meta name=\\"mobile-web-app-capable\\" content=\\"yes\\"><meta name=\\"theme-color\\" content=\\"#0f172a\\"><title>MAINT & EQUIP PRO</title><link rel=\\"icon\\" type=\\"image/png\\" href=\\"" + generatedAppIconUrl + "\\"><link rel=\\"apple-touch-icon\\" href=\\"" + generatedAppIconUrl + "\\"><style>body{margin:0;background:radial-gradient(circle at top,#1e3a8a,#0f172a 70%);color:#fff;font-family:sans-serif;display:flex;flex-direction:column;align-items:center;justify-content:center;height:100vh;text-align:center;padding:20px;box-sizing:border-box;}img{width:116px;height:116px;border-radius:26px;box-shadow:0 18px 40px rgba(0,0,0,0.6);border:2.5px solid #38bdf8;margin-bottom:16px;}a.btn{display:inline-block;margin-top:16px;padding:14px 28px;border-radius:14px;background:linear-gradient(90deg,#10b981,#06b6d4);color:#0f172a;font-weight:900;text-decoration:none;font-size:15px;box-shadow:0 10px 25px rgba(16,185,129,0.4);}</style></head><body><img src=\\"" + generatedAppIconUrl + "\\"><h2>MAINT & EQUIP PRO</h2><p style=\\"color:#cbd5e1;font-size:13px;\\">ทางลัดเข้าสู่ระบบเบิก-จ่ายอุปกรณ์ช่าง & รายงานซ่อมบำรุง A4</p><a class=\\"btn\\" href=\\"" + targetUrl + "\\">แตะเพื่อเข้าสู่แอปเต็มจอ</a><script>if(\\"" + targetUrl + "\\".indexOf(\\"http\\")===0){setTimeout(function(){window.location.replace(\\"" + targetUrl + "\\");},350);}<\\/script></body></html>";',
    '  const dataUri = "data:text/html;charset=utf-8," + encodeURIComponent(html);',
    '  triggerSafeDownload(dataUri, "MAINT_EQUIP_PRO_APP.html");',
    '  showToast("สร้างทางลัดแอป (MAINT_EQUIP_PRO_APP.html) ลงเครื่องอัตโนมัติแล้ว!");',
    '}',
    'function launchAndroidChromeInstallIntent() {',
    '  const targetUrl = getBestInstallUrl();',
    '  if (/Android/i.test(navigator.userAgent) && targetUrl.indexOf("http") === 0) {',
    '    const noProto = targetUrl.replace(/^https?:\\/\\//i, "");',
    '    const intentUrl = "intent://" + noProto + "#Intent;scheme=https;package=com.android.chrome;end";',
    '    window.location.href = intentUrl;',
    '    showToast("กำลังเปิดใน Chrome เพื่อสร้างไอคอนลงหน้าจอหลัก Android...");',
    '    return;',
    '  }',
    '  downloadMobileAppLauncherFile();',
    '  window.open(targetUrl, "_blank");',
    '}',
    'function installIosWebClipProfile() {',
    '  const targetUrl = getBestInstallUrl();',
    '  const base64Icon = generatedAppIconUrl.replace(/^data:image\\/png;base64,/, "");',
    '  const uuid1 = "A4MAINT-1111-2222-3333-" + Date.now();',
    '  const uuid2 = "A4MAINT-4444-5555-6666-" + Date.now();',
    '  const mobileConfig = "<?xml version=\\"1.0\\" encoding=\\"UTF-8\\"?><!DOCTYPE plist PUBLIC \\"-//Apple//DTD PLIST 1.0//EN\\" \\"http://www.apple.com/DTDs/PropertyList-1.0.dtd\\"><plist version=\\"1.0\\"><dict><key>PayloadContent</key><array><dict><key>FullScreen</key><true/><key>Icon</key><data>" + base64Icon + "</data><key>IsRemovable</key><true/><key>Label</key><string>MAINT PRO</string><key>PayloadDescription</key><string>Install MAINT &amp; EQUIP PRO Home Screen Icon</string><key>PayloadDisplayName</key><string>MAINT &amp; EQUIP PRO</string><key>PayloadIdentifier</key><string>com.maintpro.webclip." + uuid1 + "</string><key>PayloadType</key><string>com.apple.webClip.managed</string><key>PayloadUUID</key><string>" + uuid1 + "</string><key>PayloadVersion</key><integer>1</integer><key>Precomposed</key><true/><key>URL</key><string>" + targetUrl + "</string></dict></array><key>PayloadDisplayName</key><string>MAINT &amp; EQUIP PRO App Installer</string><key>PayloadIdentifier</key><string>com.maintpro.profile." + uuid2 + "</string><key>PayloadRemovalDisallowed</key><false/><key>PayloadType</key><string>Configuration</string><key>PayloadUUID</key><string>" + uuid2 + "</string><key>PayloadVersion</key><integer>1</integer></dict></plist>";',
    '  const dataUri = "data:application/x-apple-aspen-config;charset=utf-8," + encodeURIComponent(mobileConfig);',
    '  triggerSafeDownload(dataUri, "MAINT_PRO_iOS_Installer.mobileconfig");',
    '  showToast("ดาวน์โหลดโปรไฟล์ไอคอน iPhone แล้ว! เปิดแอป การตั้งค่า (Settings) > ดาวน์โหลดโปรไฟล์แล้ว เพื่อกดติดตั้งลงหน้าจอโฮมได้เลย");',
    '}',
    'function downloadDesktopUrlShortcut() {',
    '  const targetUrl = getBestInstallUrl();',
    '  const content = "[InternetShortcut]\\r\\nURL=" + targetUrl + "\\r\\nIconIndex=0\\r\\n";',
    '  const dataUri = "data:application/octet-stream;charset=utf-8," + encodeURIComponent(content);',
    '  triggerSafeDownload(dataUri, "MAINT_EQUIP_PRO_ทางลัดแอป.url");',
    '  showToast("สร้างทางลัด MAINT_EQUIP_PRO_ทางลัดแอป.url ลงเครื่องสำเร็จ!");',
    '}',
    'function activateFullScreenAppMode() {',
    '  try {',
    '    if (document.documentElement.requestFullscreen) {',
    '      document.documentElement.requestFullscreen();',
    '      closeInstallAppModal();',
    '      markAppAsInstalledUI();',
    '      showToast("เปิดโหมดแอปมือถือเต็มจอเรียบร้อยแล้ว!");',
    '      return;',
    '    }',
    '  } catch (e) {}',
    '  window.open(getBestInstallUrl(), "_blank");',
    '}',
    'async function shareOrOpenMobileLink() {',
    '  const url = getBestInstallUrl();',
    '  if (navigator.share) {',
    '    try { await navigator.share({ title:"MAINT & EQUIP PRO", text:"เว็บแอประบบเบิก-จ่ายอุปกรณ์ช่าง & ใบรายงานซ่อมบำรุง A4", url:url }); return; } catch(e) {}',
    '  }',
    '  copyWebAppInstallLink();',
    '}',
    'function setQuickEqDate(inputId, offsetDays) {',
    '  const d = new Date(); d.setDate(d.getDate() + offsetDays);',
    '  document.getElementById(inputId).value = d.toISOString().split("T")[0];',
    '  if (inputId === "eqReturnDate" && offsetDays === -1) {',
    '    document.getElementById("eqOverallStatus").value = "คืนของเมื่อคืนเรียบร้อย";',
    '  }',
    '  showToast("ตั้งวันที่เรียบร้อยแล้ว");',
    '}',
    'function openEquipPhotoCapture(mode, itemIdx) {',
    '  cameraModeContext = mode; cameraEquipItemIdx = itemIdx || 0;',
    '  const info = document.getElementById("cameraTargetInfo");',
    '  if (mode === "EQ_CHECKOUT") info.textContent = "ถ่ายรูปหลักฐาน: ตอนเบิกจ่าย / ขึ้นของ";',
    '  else if (mode === "EQ_RETURN") info.textContent = "ถ่ายรูปหลักฐาน: ตอนนำของมาคืน (ของเมื่อคืน/วันนี้)";',
    '  else if (mode === "EQ_ITEM") info.textContent = "ถ่ายรูปอุปกรณ์ซ่อม/อะไหล่ รายการที่ #" + (itemIdx + 1);',
    '  openCameraModal(true);',
    '}',
    'function clearEquipMainPhoto(mode) {',
    '  if (mode === "EQ_CHECKOUT") { equipState.checkoutPhoto = null; } else { equipState.returnPhoto = null; }',
    '  renderEquipMainPhotosUI();',
    '}',
    'function renderEquipMainPhotosUI() {',
    '  const pOut = document.getElementById("eqCheckoutPhotoPreview"); const eOut = document.getElementById("eqCheckoutPhotoEmpty");',
    '  if (equipState.checkoutPhoto) { pOut.src = equipState.checkoutPhoto; pOut.classList.remove("hidden"); eOut.classList.add("hidden"); }',
    '  else { pOut.classList.add("hidden"); eOut.classList.remove("hidden"); }',
    '  const pRet = document.getElementById("eqReturnPhotoPreview"); const eRet = document.getElementById("eqReturnPhotoEmpty");',
    '  if (equipState.returnPhoto) { pRet.src = equipState.returnPhoto; pRet.classList.remove("hidden"); eRet.classList.add("hidden"); }',
    '  else { pRet.classList.add("hidden"); eRet.classList.remove("hidden"); }',
    '}',
    'function stampDateWatermarkOnPhoto(dataUrl, labelTitle, dateStr, personStr, cb) {',
    '  const img = new Image();',
    '  img.onload = function() {',
    '    const c = document.createElement("canvas"); c.width = img.width; c.height = img.height;',
    '    const ctx = c.getContext("2d"); ctx.drawImage(img, 0, 0);',
    '    ctx.fillStyle = "rgba(15, 23, 42, 0.78)"; ctx.fillRect(0, c.height - 68, c.width, 68);',
    '    ctx.fillStyle = "#38bdf8"; ctx.font = "bold 20px sans-serif";',
    '    ctx.fillText(labelTitle + " | วันที่: " + (dateStr || "-"), 16, c.height - 38);',
    '    ctx.fillStyle = "#ffffff"; ctx.font = "17px sans-serif";',
    '    ctx.fillText("ผู้รับผิดชอบ: " + (personStr || "-"), 16, c.height - 14);',
    '    cb(c.toDataURL("image/jpeg", 0.85));',
    '  };',
    '  img.src = dataUrl;',
    '}',
    'function switchMainAppMode(mode) {',
    '  const eqWs = document.getElementById("equipmentAppWorkspace");',
    '  const repWs = document.getElementById("a4ReportWorkspace");',
    '  const topRepBtns = document.getElementById("topReportActionBtns");',
    '  const btnEq = document.getElementById("tabBtnEquip");',
    '  const btnRep = document.getElementById("tabBtnReport");',
    '  if (mode === "EQUIPMENT") {',
    '    eqWs.classList.remove("hidden"); repWs.classList.add("hidden"); topRepBtns.classList.add("hidden");',
    '    btnEq.className = "px-3 py-1.5 rounded-lg text-xs font-bold bg-emerald-500 text-slate-950";',
    '    btnRep.className = "px-3 py-1.5 rounded-lg text-xs font-bold text-slate-300 hover:bg-slate-700";',
    '  } else {',
    '    eqWs.classList.add("hidden"); repWs.classList.remove("hidden"); topRepBtns.classList.remove("hidden"); topRepBtns.classList.add("flex");',
    '    btnRep.className = "px-3 py-1.5 rounded-lg text-xs font-bold bg-cyan-500 text-slate-950";',
    '    btnEq.className = "px-3 py-1.5 rounded-lg text-xs font-bold text-slate-300 hover:bg-slate-700";',
    '  }',
    '}',
    'function initEquipmentWorkspace() {',
    '  const today = new Date().toISOString().split("T")[0];',
    '  document.getElementById("eqDocId").value = "EQ-" + Date.now().toString().slice(-5);',
    '  document.getElementById("eqCheckoutDate").value = today;',
    '  document.getElementById("eqReturnDate").value = "";',
    '  document.getElementById("eqRequester").value = "ช่างสมชาย (หัวหน้าทีมซ่อมบำรุง)";',
    '  document.getElementById("eqIssuer").value = "เจ้าหน้าที่สโตร์กลาง";',
    '  document.getElementById("eqReturner").value = "";',
    '  document.getElementById("eqMachineSite").value = "MCH-BLW-014";',
    '  equipState.currentItems = [',
    '    { code:"TL-001", category:"อุปกรณ์ซ่อมบำรุง", name:"เครื่องตั้งศูนย์เลเซอร์ (Laser Alignment Kit)", qtyOut:1, qtyReturn:0, unit:"ชุด", itemStatus:"กำลังใช้งานหน้างาน", remark:"เช็คแบตเตอรี่เต็ม", photo:null },',
    '    { code:"TL-002", category:"อุปกรณ์ซ่อมบำรุง", name:"เครื่องวัดความสั่นสะเทือน & อุณหภูมิ (Vib/Temp Meter)", qtyOut:1, qtyReturn:0, unit:"เครื่อง", itemStatus:"กำลังใช้งานหน้างาน", remark:"พร้อมหัววัดแม่เหล็ก", photo:null },',
    '    { code:"SP-104", category:"อะไหล่ช่าง", name:"ตลับลูกปืน SKF 6310-2Z & สายพานร่อง V", qtyOut:2, qtyReturn:0, unit:"ชุด", itemStatus:"อะไหล่ใช้เปลี่ยนหน้างาน", remark:"เบิกเปลี่ยนมอเตอร์โบลเวอร์", photo:null }',
    '  ];',
    '  equipState.records = [',
    '    { docId:"EQ-88401", checkoutDate:today, returnDate:today, overallStatus:"นำมาคืนครบแล้ว", machineOrSite:"MCH-BLW-012", requesterName:"ช่างวีระพล", issuerName:"สโตร์คลังช่าง", returnerName:"ช่างวีระพล", note:"คืนอุปกรณ์ครบสภาพสมบูรณ์", items:[{ code:"TL-005", category:"อุปกรณ์ซ่อมบำรุง", name:"รอกโซ่ 2 ตัน & ชุดประแจปอนด์", qtyOut:2, qtyReturn:2, unit:"ชุด", itemStatus:"คืนแล้ว (สภาพปกติ)", remark:"ครบถ้วน" }] },',
    '    { docId:"EQ-88390", checkoutDate:today, returnDate:"-", overallStatus:"กำลังเบิกขึ้นของเข้าทำงาน", machineOrSite:"MCH-BLW-014", requesterName:"ช่างสมชาย", issuerName:"สโตร์คลังช่าง", returnerName:"-", note:"เบิกขึ้นของเข้าหน้างาน Overhaul", items:[{ code:"TL-001", category:"อุปกรณ์ซ่อมบำรุง", name:"เครื่องตั้งศูนย์เลเซอร์ (Laser Alignment)", qtyOut:1, qtyReturn:0, unit:"ชุด", itemStatus:"กำลังใช้งานหน้างาน", remark:"ใช้งานหน้างาน" }] },',
    '    { docId:"EQ-88375", checkoutDate:today, returnDate:today, overallStatus:"อุปกรณ์ชำรุด / ส่งซ่อม", machineOrSite:"MCH-MOT-009", requesterName:"ช่างอนุชา", issuerName:"สโตร์คลังช่าง", returnerName:"ช่างอนุชา", note:"แม่แรงไฮดรอลิกซีลรั่ว ส่งซ่อม", items:[{ code:"TL-009", category:"อุปกรณ์ซ่อมบำรุง", name:"แม่แรงไฮดรอลิก 10 ตัน", qtyOut:1, qtyReturn:1, unit:"ตัว", itemStatus:"ชำรุด / ส่งซ่อม", remark:"ส่งซ่อมแผนกไฮดรอลิก" }] }',
    '  ];',
    '  renderEquipMainPhotosUI(); renderEquipItemsTable(); renderEquipHistoryTable();',
    '}',
    'function openEqStatusModal(mode) {',
    '  const all = equipState.records || [];',
    '  const filtered = all.filter(function(r){',
    '    if (mode === "ACTIVE_OUT") return r.overallStatus.indexOf("กำลังเบิก") !== -1 || r.overallStatus.indexOf("ค้างคืน") !== -1 || r.overallStatus.indexOf("รอตรวจสอบ") !== -1;',
    '    if (mode === "RETURNED") return r.overallStatus.indexOf("คืนครบ") !== -1 || r.overallStatus.indexOf("คืนของเมื่อคืน") !== -1;',
    '    if (mode === "REPAIR") return r.overallStatus.indexOf("ชำรุด") !== -1 || r.overallStatus.indexOf("ส่งซ่อม") !== -1;',
    '    return true;',
    '  });',
    '  const titles = { ALL:"ใบเบิก-จ่ายทั้งหมด", ACTIVE_OUT:"กำลังเบิกขึ้นของเข้าทำงาน (ยังไม่คืน)", RETURNED:"นำของมาคืนครบแล้ว", REPAIR:"อุปกรณ์ซ่อมชำรุด / รอซ่อม" };',
    '  document.getElementById("eqStatusModalTitle").textContent = "สถานะ: " + (titles[mode] || "ทั้งหมด") + " (" + filtered.length + " รายการ)";',
    '  const body = document.getElementById("eqStatusModalBody");',
    '  if (!filtered.length) { body.innerHTML = "<div class=\\"p-8 text-center text-slate-400 text-xs\\">ไม่มีรายการในสถานะนี้</div>"; }',
    '  else {',
    '    body.innerHTML = filtered.map(function(r){',
    '      const itemsHtml = (r.items || []).map(function(it){ return "<span class=\\"inline-block bg-slate-900 border border-slate-700 rounded px-2 py-0.5 text-[11px] mr-1 mb-1\\"><b>" + escapeHtml(it.name) + "</b> (เบิก " + it.qtyOut + "/คืน " + it.qtyReturn + " • " + escapeHtml(it.itemStatus) + ")</span>"; }).join("");',
    '      return "<div class=\\"bg-slate-800 border border-slate-700 rounded-xl p-3 space-y-1.5\\"><div class=\\"flex justify-between items-center\\"><span class=\\"font-bold text-cyan-300 text-xs\\">" + escapeHtml(r.docId) + " | " + escapeHtml(r.machineOrSite) + "</span><button onclick=\\"inspectEquipRecord(\'" + escapeHtml(r.docId) + "\');document.getElementById(\'eqStatusModal\').classList.add(\'hidden\');\\" class=\\"px-3 py-1 rounded bg-cyan-500 text-slate-950 text-xs font-extrabold\\">ดึงข้อมูลขึ้นมาดู / แก้ไข</button></div><div class=\\"text-[11px] text-slate-300\\">วันที่เบิก: <b>" + escapeHtml(r.checkoutDate) + "</b> | วันที่คืน: <b>" + escapeHtml(r.returnDate||"-") + "</b> | ผู้เบิก: <b>" + escapeHtml(r.requesterName) + "</b> | ผู้คืน: <b>" + escapeHtml(r.returnerName||"-") + "</b> | สถานะ: <b class=\\"text-amber-300\\">" + escapeHtml(r.overallStatus) + "</b></div><div>" + itemsHtml + "</div></div>";',
    '    }).join("");',
    '  }',
    '  document.getElementById("eqStatusModal").classList.remove("hidden");',
    '}',
    'function renderEquipItemsTable() {',
    '  const tb = document.getElementById("eqItemsTableBody");',
    '  tb.innerHTML = equipState.currentItems.map(function(it, idx){',
    '    const thumb = it.photo ? "<img src=\\"" + it.photo + "\\" class=\\"w-8 h-8 rounded object-cover border border-cyan-400\\">" : "<i class=\\"fa-solid fa-camera text-slate-400\\"></i>";',
    '    return "<tr>" +',
    '      "<td class=\\"p-1.5 text-slate-400 font-mono\\">" + (idx+1) + "</td>" +',
    '      "<td class=\\"p-1.5\\"><button onclick=\\"openEquipPhotoCapture(\'EQ_ITEM\'," + idx + ")\\" class=\\"w-9 h-9 rounded bg-slate-900 border border-slate-700 flex items-center justify-center\\">" + thumb + "</button></td>" +',
    '      "<td class=\\"p-1.5\\"><select onchange=\\"updateEqItem(" + idx + ",\'category\',this.value)\\" class=\\"bg-slate-900 border border-slate-700 rounded px-1.5 py-1 text-[11px] text-white\\"><option value=\\"อุปกรณ์ซ่อมบำรุง\\" " + (it.category==="อุปกรณ์ซ่อมบำรุง"?"selected":"") + ">อุปกรณ์ซ่อมบำรุง</option><option value=\\"เครื่องมือช่าง\\" " + (it.category==="เครื่องมือช่าง"?"selected":"") + ">เครื่องมือช่าง</option><option value=\\"อะไหล่ช่าง\\" " + (it.category==="อะไหล่ช่าง"?"selected":"") + ">อะไหล่ช่าง</option></select></td>" +',
    '      "<td class=\\"p-1.5\\"><input type=\\"text\\" value=\\"" + escapeHtml(it.name) + "\\" oninput=\\"updateEqItem(" + idx + ",\'name\',this.value)\\" class=\\"w-full bg-slate-900 border border-slate-700 rounded px-2 py-1 text-xs text-white font-semibold\\"></td>" +',
    '      "<td class=\\"p-1.5\\"><input type=\\"number\\" value=\\"" + (it.qtyOut||1) + "\\" oninput=\\"updateEqItem(" + idx + ",\'qtyOut\',this.value)\\" class=\\"w-14 bg-slate-900 border border-slate-700 rounded px-1 py-1 text-xs text-amber-300 text-center font-bold\\"></td>" +',
    '      "<td class=\\"p-1.5\\"><input type=\\"number\\" value=\\"" + (it.qtyReturn||0) + "\\" oninput=\\"updateEqItem(" + idx + ",\'qtyReturn\',this.value)\\" class=\\"w-14 bg-slate-900 border border-slate-700 rounded px-1 py-1 text-xs text-emerald-300 text-center font-bold\\"></td>" +',
    '      "<td class=\\"p-1.5\\"><input type=\\"text\\" value=\\"" + escapeHtml(it.unit||"ชุด") + "\\" oninput=\\"updateEqItem(" + idx + ",\'unit\',this.value)\\" class=\\"w-14 bg-slate-900 border border-slate-700 rounded px-1 py-1 text-xs text-white text-center\\"></td>" +',
    '      "<td class=\\"p-1.5\\"><select onchange=\\"updateEqItem(" + idx + ",\'itemStatus\',this.value)\\" class=\\"bg-slate-900 border border-slate-700 rounded px-1.5 py-1 text-[11px] text-cyan-300 font-semibold\\"><option value=\\"กำลังใช้งานหน้างาน\\" " + (it.itemStatus==="กำลังใช้งานหน้างาน"?"selected":"") + ">กำลังใช้งานหน้างาน</option><option value=\\"คืนแล้ว (สภาพปกติ)\\" " + (it.itemStatus==="คืนแล้ว (สภาพปกติ)"?"selected":"") + ">คืนแล้ว (สภาพปกติ)</option><option value=\\"อะไหล่ใช้เปลี่ยนหน้างาน\\" " + (it.itemStatus==="อะไหล่ใช้เปลี่ยนหน้างาน"?"selected":"") + ">อะไหล่ใช้เปลี่ยนหน้างาน</option><option value=\\"ชำรุด / ส่งซ่อม\\" " + (it.itemStatus==="ชำรุด / ส่งซ่อม"?"selected":"") + ">ชำรุด / ส่งซ่อม</option></select></td>" +',
    '      "<td class=\\"p-1.5\\"><input type=\\"text\\" value=\\"" + escapeHtml(it.remark||"") + "\\" oninput=\\"updateEqItem(" + idx + ",\'remark\',this.value)\\" class=\\"w-full bg-slate-900 border border-slate-700 rounded px-2 py-1 text-xs text-slate-300\\"></td>" +',
    '      "<td class=\\"p-1.5\\"><button onclick=\\"removeEqItem(" + idx + ")\\" class=\\"text-rose-400 hover:text-rose-300 px-1\\"><i class=\\"fa-solid fa-trash\\"></i></button></td>" +',
    '    "</tr>";',
    '  }).join("");',
    '}',
    'function updateEqItem(idx, key, val) { if (equipState.currentItems[idx]) equipState.currentItems[idx][key] = val; }',
    'function removeEqItem(idx) { if (equipState.currentItems.length > 1) { equipState.currentItems.splice(idx,1); renderEquipItemsTable(); } }',
    'function addBlankEquipItemRow() {',
    '  equipState.currentItems.push({ code:"EQ-"+(equipState.currentItems.length+1), category:"อุปกรณ์ซ่อมบำรุง", name:"", qtyOut:1, qtyReturn:0, unit:"ชุด", itemStatus:"กำลังใช้งานหน้างาน", remark:"" });',
    '  renderEquipItemsTable();',
    '}',
    'function addQuickPresetEquipment() {',
    '  equipState.currentItems.push(',
    '    { code:"TL-101", category:"อุปกรณ์ซ่อมบำรุง", name:"เครื่องดูดลูกปืนไฮดรอลิก (Hydraulic Puller)", qtyOut:1, qtyReturn:0, unit:"ชุด", itemStatus:"กำลังใช้งานหน้างาน", remark:"ชุดใหญ่" },',
    '    { code:"TL-102", category:"เครื่องมือช่าง", name:"ชุดบล็อกลม & ประแจปอนด์มาตรฐาน", qtyOut:1, qtyReturn:0, unit:"กล่อง", itemStatus:"กำลังใช้งานหน้างาน", remark:"ครบชุด" },',
    '    { code:"SP-201", category:"อะไหล่ช่าง", name:"จาระบีทนความร้อนสูง SKF LGHP 2", qtyOut:1, qtyReturn:0, unit:"กระป๋อง", itemStatus:"อะไหล่ใช้เปลี่ยนหน้างาน", remark:"อัดลูกปืนมอเตอร์" }',
    '  );',
    '  renderEquipItemsTable(); showToast("เพิ่มชุดอุปกรณ์ซ่อมบำรุงและอะไหล่ช่างเรียบร้อยแล้ว");',
    '}',
    'function resetEquipForm() {',
    '  document.getElementById("eqDocId").value = "EQ-" + Date.now().toString().slice(-5);',
    '  document.getElementById("eqCheckoutDate").value = new Date().toISOString().split("T")[0];',
    '  document.getElementById("eqReturnDate").value = "";',
    '  document.getElementById("eqOverallStatus").value = "กำลังเบิกขึ้นของเข้าทำงาน";',
    '  document.getElementById("eqReturner").value = "";',
    '  equipState.checkoutPhoto = null; equipState.returnPhoto = null;',
    '  equipState.currentItems = [{ code:"TL-01", category:"อุปกรณ์ซ่อมบำรุง", name:"", qtyOut:1, qtyReturn:0, unit:"ชุด", itemStatus:"กำลังใช้งานหน้างาน", remark:"", photo:null }];',
    '  renderEquipMainPhotosUI(); renderEquipItemsTable(); showToast("เริ่มใบเบิกอุปกรณ์ใหม่");',
    '}',
    'function markCurrentEquipReturnedToday() {',
    '  const today = new Date().toISOString().split("T")[0];',
    '  if (!document.getElementById("eqReturnDate").value) document.getElementById("eqReturnDate").value = today;',
    '  document.getElementById("eqOverallStatus").value = "นำมาคืนครบแล้ว";',
    '  if (!document.getElementById("eqReturner").value) document.getElementById("eqReturner").value = document.getElementById("eqRequester").value;',
    '  equipState.currentItems.forEach(function(it){ if (it.category !== "อะไหล่ช่าง") { it.qtyReturn = it.qtyOut; it.itemStatus = "คืนแล้ว (สภาพปกติ)"; } });',
    '  renderEquipItemsTable(); saveEquipRecord();',
    '}',
    'function saveEquipRecord() {',
    '  document.getElementById("eqOverallStatus").value = "รอตรวจสอบ";',
    '  const rec = {',
    '    recordType: "EQUIPMENT_REQUISITION",',
    '    docId: document.getElementById("eqDocId").value.trim() || ("EQ-" + Date.now().toString().slice(-5)),',
    '    checkoutDate: document.getElementById("eqCheckoutDate").value || new Date().toISOString().split("T")[0],',
    '    returnDate: document.getElementById("eqReturnDate").value || "-",',
    '    overallStatus: "รอตรวจสอบ",',
    '    machineOrSite: document.getElementById("eqMachineSite").value || "-",',
    '    requesterName: document.getElementById("eqRequester").value || "-",',
    '    issuerName: document.getElementById("eqIssuer").value || "-",',
    '    returnerName: document.getElementById("eqReturner").value || "-",',
    '    receiverName: document.getElementById("eqIssuer").value || "-",',
    '    note: document.getElementById("eqNote").value || "-",',
    '    checkoutPhoto: equipState.checkoutPhoto || null,',
    '    returnPhoto: equipState.returnPhoto || null,',
    '    items: JSON.parse(JSON.stringify(equipState.currentItems))',
    '  };',
    '  const existIdx = equipState.records.findIndex(function(r){ return r.docId === rec.docId; });',
    '  if (existIdx !== -1) equipState.records[existIdx] = rec; else equipState.records.unshift(rec);',
    '  renderEquipHistoryTable();',
    '  if (typeof google !== "undefined" && google.script && google.script.run) {',
    '    google.script.run.withSuccessHandler(function(res){',
    '      if (res && res.records) { equipState.records = res.records; renderEquipHistoryTable(); }',
    '      showToast("บันทึกใบเบิก-คืนอุปกรณ์ (" + rec.docId + ") ลง Google Sheets สำเร็จ!");',
    '    }).saveEquipmentFromClient(rec);',
    '  } else { showToast("บันทึกใบเบิก-คืนอุปกรณ์ (" + rec.docId + ") สำเร็จ!"); }',
    '}',
    'function loadEquipHistoryFromServer() {',
    '  if (typeof google !== "undefined" && google.script && google.script.run) {',
    '    google.script.run.withSuccessHandler(function(res){',
    '      if (res && Array.isArray(res.records) && res.records.length) { equipState.records = res.records; renderEquipHistoryTable(); }',
    '    }).getEquipmentRecordsFromClient();',
    '  }',
    '}',
    'function renderEquipHistoryTable() {',
    '  const q = (document.getElementById("eqSearchInput").value || "").toLowerCase();',
    '  const st = document.getElementById("eqFilterStatus").value;',
    '  const all = equipState.records || [];',
    '  document.getElementById("eqStatTotal").textContent = all.length;',
    '  document.getElementById("eqStatOut").textContent = all.filter(function(r){ return r.overallStatus.indexOf("กำลังเบิก") !== -1 || r.overallStatus.indexOf("ค้างคืน") !== -1; }).length;',
    '  document.getElementById("eqStatReturned").textContent = all.filter(function(r){ return r.overallStatus.indexOf("คืนครบ") !== -1; }).length;',
    '  document.getElementById("eqStatRepair").textContent = all.filter(function(r){ return r.overallStatus.indexOf("ชำรุด") !== -1; }).length;',
    '  const filtered = all.filter(function(r){',
    '    if (st !== "ALL" && r.overallStatus !== st) return false;',
    '    const text = [r.docId, r.checkoutDate, r.returnDate, r.requesterName, r.returnerName, r.machineOrSite, JSON.stringify(r.items||[])].join(" ").toLowerCase();',
    '    return text.indexOf(q) !== -1;',
    '  });',
    '  const tb = document.getElementById("eqHistoryTableBody");',
    '  if (!filtered.length) { tb.innerHTML = "<tr><td colspan=\\"4\\" class=\\"p-4 text-center text-slate-400\\">ไม่พบรายการที่ค้นหา</td></tr>"; return; }',
    '  tb.innerHTML = filtered.map(function(r){',
    '    const badgeCls = r.overallStatus.indexOf("คืนครบ") !== -1 ? "bg-emerald-500/20 text-emerald-300 border-emerald-500/40" : (r.overallStatus.indexOf("ชำรุด") !== -1 ? "bg-rose-500/20 text-rose-300 border-rose-500/40" : "bg-amber-500/20 text-amber-300 border-amber-500/40");',
    '    return "<tr class=\\"hover:bg-slate-900/70\\">" +',
    '      "<td class=\\"p-2\\"><div class=\\"font-bold text-cyan-300\\">" + escapeHtml(r.docId) + " (" + escapeHtml(r.machineOrSite) + ")</div><div class=\\"text-[10px] text-slate-400\\">เบิก: " + escapeHtml(r.checkoutDate) + " | คืน: " + escapeHtml(r.returnDate||"-") + "</div></td>" +',
    '      "<td class=\\"p-2\\"><div class=\\"text-white font-semibold\\">ผู้เบิก: " + escapeHtml(r.requesterName) + "</div><div class=\\"text-[10px] text-emerald-300\\">ผู้คืน: " + escapeHtml(r.returnerName||"-") + "</div></td>" +',
    '      "<td class=\\"p-2\\"><span class=\\"px-2 py-0.5 rounded-full text-[10px] font-bold border " + badgeCls + "\\">" + escapeHtml(r.overallStatus) + "</span><div class=\\"text-[10px] text-slate-400 mt-0.5\\">รวม " + ((r.items||[]).length) + " รายการ</div></td>" +',
    '      "<td class=\\"p-2 text-right\\"><button onclick=\\"inspectEquipRecord(\'" + escapeHtml(r.docId) + "\')\\" class=\\"px-2.5 py-1 rounded bg-cyan-600 hover:bg-cyan-500 text-white text-[11px] font-bold\\">ดึงข้อมูลดู/รับคืน</button></td>" +',
    '    "</tr>";',
    '  }).join("");',
    '}',
    'function inspectEquipRecord(docId) {',
    '  const rec = equipState.records.find(function(r){ return r.docId === docId; }); if (!rec) return;',
    '  document.getElementById("eqDocId").value = rec.docId;',
    '  document.getElementById("eqCheckoutDate").value = rec.checkoutDate !== "-" ? rec.checkoutDate : "";',
    '  document.getElementById("eqReturnDate").value = rec.returnDate !== "-" ? rec.returnDate : "";',
    '  document.getElementById("eqOverallStatus").value = rec.overallStatus;',
    '  document.getElementById("eqMachineSite").value = rec.machineOrSite || "";',
    '  document.getElementById("eqRequester").value = rec.requesterName || "";',
    '  document.getElementById("eqIssuer").value = rec.issuerName || "";',
    '  document.getElementById("eqReturner").value = rec.returnerName !== "-" ? (rec.returnerName || "") : "";',
    '  document.getElementById("eqNote").value = rec.note !== "-" ? (rec.note || "") : "";',
    '  equipState.checkoutPhoto = rec.checkoutPhoto || null;',
    '  equipState.returnPhoto = rec.returnPhoto || null;',
    '  equipState.currentItems = Array.isArray(rec.items) && rec.items.length ? JSON.parse(JSON.stringify(rec.items)) : [];',
    '  renderEquipMainPhotosUI(); renderEquipItemsTable(); showToast("ดึงข้อมูลใบเบิก " + rec.docId + " ขึ้นมาแสดงแล้ว");',
    '}',
    'function printEquipSlipWithPhotos() {',
    '  const docId = document.getElementById("eqDocId").value || "-";',
    '  const outDate = document.getElementById("eqCheckoutDate").value || "-";',
    '  const retDate = document.getElementById("eqReturnDate").value || "-";',
    '  const req = document.getElementById("eqRequester").value || "-";',
    '  const ret = document.getElementById("eqReturner").value || "-";',
    '  const w = window.open("", "_blank", "width=900,height=750"); if (!w) return;',
    '  const rows = equipState.currentItems.map(function(it,i){ return "<tr><td style=\\"border:1px solid #333;padding:5px;text-align:center\\">"+(i+1)+"</td><td style=\\"border:1px solid #333;padding:5px\\">"+escapeHtml(it.category)+"</td><td style=\\"border:1px solid #333;padding:5px;font-weight:bold\\">"+escapeHtml(it.name)+"</td><td style=\\"border:1px solid #333;padding:5px;text-align:center\\">"+(it.qtyOut||0)+"</td><td style=\\"border:1px solid #333;padding:5px;text-align:center\\">"+(it.qtyReturn||0)+"</td><td style=\\"border:1px solid #333;padding:5px\\">"+escapeHtml(it.itemStatus)+"</td></tr>"; }).join("");',
    '  const outImg = equipState.checkoutPhoto ? "<img src=\\""+equipState.checkoutPhoto+"\\" style=\\"width:180px;height:180px;object-fit:cover;border:1px solid #333;\\">" : "<div style=\\"width:180px;height:180px;border:1px dashed #999;line-height:180px;text-align:center\\">ไม่มีรูปตอนเบิก</div>";',
    '  const retImg = equipState.returnPhoto ? "<img src=\\""+equipState.returnPhoto+"\\" style=\\"width:180px;height:180px;object-fit:cover;border:1px solid #333;\\">" : "<div style=\\"width:180px;height:180px;border:1px dashed #999;line-height:180px;text-align:center\\">ไม่มีรูปตอนคืน</div>";',
    '  w.document.write("<html><head><title>ใบเบิก-คืนอุปกรณ์ "+escapeHtml(docId)+"</title></head><body style=\\"font-family:sans-serif;padding:20px\\"><h2>ใบเบิก-จ่ายและรับคืนอุปกรณ์ซ่อมบำรุง / อะไหล่ช่าง ("+escapeHtml(docId)+")</h2><p><b>วันที่เบิกจ่าย:</b> "+escapeHtml(outDate)+" (ผู้เบิก: "+escapeHtml(req)+") | <b>วันที่นำมาคืน:</b> "+escapeHtml(retDate)+" (ผู้คืน: "+escapeHtml(ret)+")</p><table style=\\"width:100%;border-collapse:collapse;margin-bottom:16px\\"><tr style=\\"background:#eee\\"><th style=\\"border:1px solid #333;padding:5px\\">#</th><th style=\\"border:1px solid #333;padding:5px\\">หมวดหมู่</th><th style=\\"border:1px solid #333;padding:5px\\">รายการอุปกรณ์</th><th style=\\"border:1px solid #333;padding:5px\\">เบิก</th><th style=\\"border:1px solid #333;padding:5px\\">คืน</th><th style=\\"border:1px solid #333;padding:5px\\">สถานะ</th></tr>"+rows+"</table><div style=\\"display:flex;gap:30px\\"><div><b>รูปหลักฐานวันเบิกจ่าย:</b><br>"+outImg+"</div><div><b>รูปหลักฐานวันนำมาคืน:</b><br>"+retImg+"</div></div><script>window.onload=function(){window.print();}<\\/script></body></html>");',
    '  w.document.close();',
    '}',
    'function computeSafeFrameStyle(gridCount, orientation, secCode) {',
    '  const reqCm = parseFloat(appState.frameSize) || 6; const isL = orientation === "landscape";',
    '  const hasExtraTopPanel = (secCode === "SEC_6" || secCode === "SEC_7");',
    '  let cols = 2, rows = 2;',
    '  if (gridCount === 6) { cols = isL ? 3 : 2; rows = isL ? 2 : 3; }',
    '  else if (gridCount === 8) { cols = isL ? 4 : 2; rows = isL ? 2 : 4; }',
    '  else if (gridCount === 12) { cols = isL ? 4 : 3; rows = isL ? 3 : 4; }',
    '  const availH = isL ? (hasExtraTopPanel ? 10.5 : 13.8) : (hasExtraTopPanel ? 17.2 : 21.0);',
    '  const maxW = ((isL ? 27.2 : 18.6) / cols) - 0.5;',
    '  const maxH = (availH / rows) - (rows >= 4 ? 1.15 : 1.45);',
    '  const safeCm = Math.max(2.4, Math.min(reqCm, maxW, maxH)).toFixed(2);',
    '  return { cols:cols, rows:rows, safeCm:safeCm, reqCm:reqCm };',
    '}',
    'function renderAll() {',
    '  document.documentElement.style.setProperty("--frame-size", appState.frameSize);',
    '  document.getElementById("dynamicPageStyle").innerHTML = "@page { size: A4 " + appState.orientation + "; margin: 0; }";',
    '  const cur = appState.pages[appState.activePageIndex];',
    '  document.getElementById("currentPageIndicator").textContent = "หน้า " + (appState.activePageIndex+1) + "/" + appState.pages.length;',
    '  document.getElementById("pageSectionSelect").value = cur.sectionCode;',
    '  document.getElementById("pageTabsContainer").innerHTML = appState.pages.map(function(p, idx) {',
    '    const sec = SECTION_DEFINITIONS[p.sectionCode] || SECTION_DEFINITIONS.SEC_1; const active = idx === appState.activePageIndex;',
    '    return "<button onclick=\\"selectActivePage(" + idx + ")\\" class=\\"px-2.5 py-1 rounded text-xs font-semibold shrink-0 " + (active ? "bg-cyan-500 text-slate-950" : "bg-slate-800 text-slate-300") + "\\">หน้า " + (idx+1) + " (หัวข้อที่ " + sec.num + " • " + p.gridCount + " รูป)</button>";',
    '  }).join("");',
    '  renderA4Pages();',
    '}',
    'function renderA4Pages() {',
    '  const isL = appState.orientation === "landscape";',
    '  document.getElementById("a4PagesContainer").innerHTML = appState.pages.map(function(page, pIdx) {',
    '    const sec = SECTION_DEFINITIONS[page.sectionCode] || SECTION_DEFINITIONS.SEC_1; const dim = computeSafeFrameStyle(page.gridCount, appState.orientation, page.sectionCode);',
    '    let dedicatedSectionPanel = "";',
    '    if (page.sectionCode === "SEC_6") {',
    '      dedicatedSectionPanel = "<div class=\\"mb-1.5 border-2 border-rose-700 rounded-lg overflow-hidden bg-rose-50/60 shrink-0\\">" +',
    '        "<div class=\\"bg-rose-800 text-white px-3 py-1 text-xs font-bold flex justify-between items-center\\"><span>ตารางบันทึกค่าหัวข้อที่ 6 : ความร้อน Temperature (°C) 3 ค่า และ ความสั่นสะเทือน Vibration (mm/s) 3 ค่า</span><span>SECTION 6 ONLY</span></div>" +',
    '        "<div class=\\"p-2 grid grid-cols-2 gap-2 text-xs\\">" +',
    '          "<div class=\\"bg-white border border-orange-300 rounded p-1.5\\"><div class=\\"font-bold text-orange-800 border-b border-orange-200 pb-0.5 mb-1\\">6.1 ค่าความร้อน Temperature (°C) - 3 ค่า</div><div class=\\"grid grid-cols-3 gap-1 text-center\\"><div class=\\"bg-orange-50 rounded p-1\\"><div class=\\"text-[9px] text-slate-500\\">จุดที่ 1 (DE)</div><div class=\\"font-extrabold text-slate-900\\">" + escapeHtml(appState.tempValues[0]) + " °C</div></div><div class=\\"bg-orange-50 rounded p-1\\"><div class=\\"text-[9px] text-slate-500\\">จุดที่ 2 (NDE)</div><div class=\\"font-extrabold text-slate-900\\">" + escapeHtml(appState.tempValues[1]) + " °C</div></div><div class=\\"bg-orange-50 rounded p-1\\"><div class=\\"text-[9px] text-slate-500\\">จุดที่ 3 (Blower)</div><div class=\\"font-extrabold text-slate-900\\">" + escapeHtml(appState.tempValues[2]) + " °C</div></div></div></div>" +',
    '          "<div class=\\"bg-white border border-sky-300 rounded p-1.5\\"><div class=\\"font-bold text-sky-800 border-b border-sky-200 pb-0.5 mb-1\\">6.2 ค่าความสั่นสะเทือน Vibration (mm/s) - 3 ค่า</div><div class=\\"grid grid-cols-3 gap-1 text-center\\"><div class=\\"bg-sky-50 rounded p-1\\"><div class=\\"text-[9px] text-slate-500\\">แนวราบ (H)</div><div class=\\"font-extrabold text-slate-900\\">" + escapeHtml(appState.vibValues[0]) + " mm/s</div></div><div class=\\"bg-sky-50 rounded p-1\\"><div class=\\"text-[9px] text-slate-500\\">แนวดิ่ง (V)</div><div class=\\"font-extrabold text-slate-900\\">" + escapeHtml(appState.vibValues[1]) + " mm/s</div></div><div class=\\"bg-sky-50 rounded p-1\\"><div class=\\"text-[9px] text-slate-500\\">แนวแกน (A)</div><div class=\\"font-extrabold text-slate-900\\">" + escapeHtml(appState.vibValues[2]) + " mm/s</div></div></div></div>" +',
    '        "</div></div>";',
    '    } else if (page.sectionCode === "SEC_7") {',
    '      dedicatedSectionPanel = "<div class=\\"mb-1.5 border-2 border-purple-800 rounded-lg overflow-hidden bg-purple-50/60 shrink-0\\">" +',
    '        "<div class=\\"bg-purple-900 text-white px-3 py-1 text-xs font-bold flex justify-between items-center\\"><span>ตารางบันทึกหัวข้อที่ 7 : ข้อเสนอแนะ Comment / รหัสแจ้งเตือน Alarm Code</span><span class=\\"bg-amber-400 text-slate-950 px-2 py-0.5 rounded font-extrabold\\">Alarm: " + escapeHtml(appState.alarmCode) + "</span></div>" +',
    '        "<div class=\\"p-2.5 bg-white text-xs text-slate-900 space-y-1\\">" +',
    '          "<div><b>รหัสแจ้งเตือน (Alarm Code):</b> <span class=\\"text-rose-700 font-extrabold\\">" + escapeHtml(appState.alarmCode) + "</span></div>" +',
    '          "<div><b>ข้อเสนอแนะ / รายละเอียดการแก้ไข (Comment):</b> " + escapeHtml(appState.comment) + "</div>" +',
    '        "</div></div>";',
    '    }',
    '    const slotsHtml = page.slots.map(function(slot, sIdx) {',
    '      const badgeBg = slot.statusTag.indexOf("ก่อนทำ") !== -1 ? "#dc2626" : (slot.statusTag.indexOf("หลังทำ") !== -1 ? "#16a34a" : "#2563eb");',
    '      const rmkHtml = (slot.remarks || []).map(function(r){ return "<div class=\\"text-[10px] text-slate-700 truncate\\">• " + escapeHtml(r) + "</div>"; }).join("");',
    '      const imgHtml = slot.imageSrc ? "<img src=\\"" + slot.imageSrc + "\\" style=\\"object-fit:cover;\\">" : "<div class=\\"text-center text-slate-400 text-[9px] p-1 pb-6\\"><i class=\\"fa-solid fa-camera text-base block mb-0.5 text-indigo-500\\"></i>กดถ่ายรูป / เพิ่มรูป<br>" + dim.reqCm + "x" + dim.reqCm + " ซม.</div>";',
    '      const innerBottomBtns = "<div class=\\"slot-action-overlay no-print absolute bottom-1 inset-x-1 z-10 flex items-center justify-center gap-1 bg-slate-900/80 py-1 px-1 rounded\\">" +',
    '        "<button type=\\"button\\" onclick=\\"event.stopPropagation();openSlotCam(" + pIdx + "," + sIdx + ")\\" class=\\"px-1.5 py-0.5 rounded bg-indigo-600 hover:bg-indigo-500 text-white text-[9px] font-bold\\"><i class=\\"fa-solid fa-camera mr-0.5\\"></i>ถ่ายรูป</button>" +',
    '        "<button type=\\"button\\" onclick=\\"event.stopPropagation();onSlotPhotoClick(" + pIdx + "," + sIdx + ")\\" class=\\"px-1.5 py-0.5 rounded bg-blue-600 hover:bg-blue-500 text-white text-[9px] font-bold\\"><i class=\\"fa-solid fa-image mr-0.5\\"></i>เพิ่มรูป</button>" +',
    '        "<button type=\\"button\\" onclick=\\"event.stopPropagation();clearSlotImg(" + pIdx + "," + sIdx + ")\\" class=\\"px-1.5 py-0.5 rounded bg-rose-600 hover:bg-rose-500 text-white text-[9px] font-bold\\"><i class=\\"fa-solid fa-trash mr-0.5\\"></i>ลบรูป</button>" +',
    '      "</div>";',
    '      return "<div class=\\"bg-transparent p-1 flex flex-col justify-between relative\\">" +',
    '        "<div class=\\"flex justify-between items-center text-[10px] font-bold px-1 mb-0.5\\"><span class=\\"truncate\\">" + escapeHtml(slot.sectionTitle) + "</span><span class=\\"px-1.5 py-0.2 rounded text-[9px] text-white\\" style=\\"background:" + badgeBg + "\\">" + escapeHtml(slot.statusTag) + "</span></div>" +',
    '        "<div class=\\"my-auto py-0.5 flex items-center justify-center\\"><div onclick=\\"openSlotCam(" + pIdx + "," + sIdx + ")\\" class=\\"photo-frame-box " + (slot.imageSrc ? "has-image" : "") + " cursor-pointer\\" style=\\"width:" + dim.safeCm + "cm;height:" + dim.safeCm + "cm;\\">" + imgHtml + innerBottomBtns + "</div></div>" +',
    '        "<div onclick=\\"openSlotEditModal(" + pIdx + "," + sIdx + ")\\" class=\\"mt-1 pt-1 border-t border-slate-200 bg-slate-50 rounded px-1.5 py-0.5 cursor-pointer\\"><div class=\\"flex justify-between text-[10px] font-bold text-slate-900\\"><span class=\\"truncate\\">หัวข้อ: " + escapeHtml(slot.subTitle) + "</span><span class=\\"no-print text-[9px] text-blue-600\\">แก้ไข</span></div>" + rmkHtml + "</div></div>";',
    '    }).join("");',
    '    return "<div class=\\"a4-sheet " + (isL ? "landscape" : "") + "\\">" +',
    '      "<div class=\\"border-b-2 border-slate-800 pb-1 mb-1 shrink-0\\"><div class=\\"flex justify-between items-center\\"><div><h2 class=\\"text-sm font-extrabold text-slate-900\\">ใบรายงานการซ่อมบำรุงและโอเวอร์ฮอลเครื่องจักร (MAINTENANCE REPORT)</h2><p class=\\"text-[10px] text-slate-600\\">เครื่องจักร: <b>" + escapeHtml(appState.machineName) + "</b> | ผู้ปฏิบัติงาน: <b>" + escapeHtml(appState.techName) + "</b></p></div>" +',
    '      "<div class=\\"flex gap-1.5 text-right\\"><div class=\\"border border-slate-400 rounded px-2 py-0.5 bg-slate-50\\"><div class=\\"text-[8px] text-slate-500\\">MACHINE CODE</div><div class=\\"text-xs font-extrabold text-blue-900\\">" + escapeHtml(appState.machineCode) + "</div></div><div class=\\"border border-slate-400 rounded px-2 py-0.5 bg-slate-50\\"><div class=\\"text-[8px] text-slate-500\\">DATE</div><div class=\\"text-xs font-bold text-slate-900\\">" + escapeHtml(appState.serviceDate) + "</div></div></div></div>" +',
    '      "<div class=\\"mt-1 rounded px-2.5 py-1 flex justify-between items-center\\" style=\\"background:" + sec.bgLight + ";border:1.5px solid " + sec.borderColor + "\\"><span class=\\"text-xs font-extrabold text-slate-900\\">สัดส่วนเฉพาะ " + escapeHtml(sec.title) + "</span><span class=\\"text-[10px] font-semibold text-slate-700\\">ตาราง " + page.gridCount + " รูป | กรอบ " + dim.reqCm + "x" + dim.reqCm + " ซม. | หน้า " + (pIdx+1) + "/" + appState.pages.length + "</span></div></div>" +',
    '      dedicatedSectionPanel +',
    '      "<div class=\\"flex-1 grid gap-2 my-0.5 overflow-hidden\\" style=\\"grid-template-columns:repeat(" + dim.cols + ",minmax(0,1fr));grid-template-rows:repeat(" + dim.rows + ",minmax(0,1fr));\\">" + slotsHtml + "</div>" +',
    '      "<div class=\\"mt-1 pt-1 border-t border-slate-300 flex justify-between items-center text-[10px] text-slate-600 shrink-0\\"><span>สัดส่วนรายงาน: <b>" + escapeHtml(sec.title) + "</b> (แยกสัดส่วนไม่ปนกับหัวข้ออื่น)</span><span>ผู้ปฏิบัติงาน: <b>" + escapeHtml(appState.techName) + "</b> | สถานะ: <b>" + escapeHtml(appState.jobStatus) + "</b></span></div></div>";',
    '  }).join("");',
    '}',
    'function syncJobData() {',
    '  appState.machineCode = document.getElementById("inputMachineCode").value || "MCH-000";',
    '  appState.serviceDate = document.getElementById("inputServiceDate").value;',
    '  appState.machineName = document.getElementById("inputMachineName").value;',
    '  appState.techName = document.getElementById("inputTechName").value;',
    '  appState.tempValues = [document.getElementById("tempVal1").value, document.getElementById("tempVal2").value, document.getElementById("tempVal3").value];',
    '  appState.vibValues = [document.getElementById("vibVal1").value, document.getElementById("vibVal2").value, document.getElementById("vibVal3").value];',
    '  appState.alarmCode = document.getElementById("inputAlarmCode").value;',
    '  appState.comment = document.getElementById("inputComment").value;',
    '  renderA4Pages();',
    '}',
    'function setOrientation(m) { appState.orientation = m; renderAll(); }',
    'function setFrameSize(s) { appState.frameSize = s; renderAll(); }',
    'function setPageGridCount(cnt) {',
    '  const p = appState.pages[appState.activePageIndex]; const sec = SECTION_DEFINITIONS[p.sectionCode];',
    '  while (p.slots.length < cnt) { const i = p.slots.length; p.slots.push({ sectionTitle:sec.title.split(" (")[0], statusTag:(i%2===0)?"ก่อนทำ (BEFORE)":"หลังทำ (AFTER)", subTitle:sec.defaultSub+" #"+(Math.floor(i/2)+1), remarks:["ตรวจสอบเรียบร้อย"], imageSrc:null }); }',
    '  if (p.slots.length > cnt) p.slots = p.slots.slice(0, cnt);',
    '  p.gridCount = cnt; renderAll();',
    '}',
    'function changeActivePageSection(code) {',
    '  const p = appState.pages[appState.activePageIndex]; p.sectionCode = code; const sec = SECTION_DEFINITIONS[code];',
    '  p.slots.forEach(function(s){ s.sectionTitle = sec.title.split(" (")[0]; }); renderAll();',
    '}',
    'function selectActivePage(i) { appState.activePageIndex = i; renderAll(); }',
    'function addNewA4Page() { appState.pages.push(createPageObject(appState.pages[appState.activePageIndex].sectionCode, 4)); appState.activePageIndex = appState.pages.length - 1; renderAll(); }',
    'function removeCurrentA4Page() { if (appState.pages.length > 1) { appState.pages.splice(appState.activePageIndex, 1); appState.activePageIndex = 0; renderAll(); } }',
    'function openOrCreateSectionPage(secCode) {',
    '  const existingIdx = appState.pages.findIndex(function(p){ return p.sectionCode === secCode; });',
    '  if (existingIdx !== -1) { appState.activePageIndex = existingIdx; }',
    '  else { appState.pages.push(createPageObject(secCode, 4)); appState.activePageIndex = appState.pages.length - 1; }',
    '  renderAll(); showToast("เปิดหน้ากระดาษเฉพาะ " + SECTION_DEFINITIONS[secCode].title);',
    '}',
    'function generateAll7SectionsPages() {',
    '  ["SEC_1","SEC_2","SEC_3","SEC_4","SEC_5","SEC_6","SEC_7"].forEach(function(c){',
    '    if (!appState.pages.some(function(p){return p.sectionCode===c;})) appState.pages.push(createPageObject(c, 4));',
    '  });',
    '  renderAll(); showToast("สร้างหน้ากระดาษแยกสัดส่วนครบทั้ง 7 หัวข้อ (1 ถึง 7) เรียบร้อยแล้ว");',
    '}',
    'function processImageFileToDataURL(file, cb) {',
    '  const reader = new FileReader(); reader.onload = function(e) {',
    '    const img = new Image(); img.onload = function() {',
    '      const canvas = document.createElement("canvas"); const max = 900; let w = img.width, h = img.height;',
    '      if (w > max || h > max) { if (w > h) { h = Math.round(h*max/w); w = max; } else { w = Math.round(w*max/h); h = max; } }',
    '      canvas.width = w; canvas.height = h; canvas.getContext("2d").drawImage(img, 0, 0, w, h); cb(canvas.toDataURL("image/jpeg", 0.85));',
    '    }; img.src = e.target.result;',
    '  }; reader.readAsDataURL(file);',
    '}',
    'function onSlotPhotoClick(pIdx, sIdx) { targetSlotRef = { pageIndex:pIdx, slotIndex:sIdx }; document.getElementById("singleSlotFileInput").click(); }',
    'function openSlotCam(pIdx, sIdx) { targetSlotRef = { pageIndex:pIdx, slotIndex:sIdx }; cameraModeContext = "REPORT"; openCameraModal(true); }',
    'function clearSlotImg(pIdx, sIdx) { if (appState.pages[pIdx] && appState.pages[pIdx].slots[sIdx]) { appState.pages[pIdx].slots[sIdx].imageSrc = null; renderA4Pages(); showToast("ลบรูปภาพแล้ว"); } }',
    'function handleSingleSlotFile(e) { const f = e.target.files[0]; if (!f) return; processImageFileToDataURL(f, function(url){ appState.pages[targetSlotRef.pageIndex].slots[targetSlotRef.slotIndex].imageSrc = url; renderA4Pages(); }); }',
    'function triggerBatchUpload() { document.getElementById("batchFileInput").click(); }',
    'function handleBatchFiles(e) {',
    '  const files = Array.from(e.target.files); if (!files.length) return; let done = 0;',
    '  files.forEach(function(f){ processImageFileToDataURL(f, function(url){ placePhotoNext(url); done++; if (done===files.length) { renderAll(); showToast("เพิ่มรูปพร้อมกัน " + files.length + " รูปสำเร็จ"); } }); });',
    '}',
    'function placePhotoNext(url) {',
    '  for (let p=appState.activePageIndex; p<appState.pages.length; p++) {',
    '    for (let s=0; s<appState.pages[p].slots.length; s++) { if (!appState.pages[p].slots[s].imageSrc) { appState.pages[p].slots[s].imageSrc = url; return; } }',
    '  }',
    '  const np = createPageObject(appState.pages[appState.activePageIndex].sectionCode, 4); np.slots[0].imageSrc = url; appState.pages.push(np); appState.activePageIndex = appState.pages.length - 1;',
    '}',
    'async function openCameraModal(forSpecificSlot) {',
    '  if (!forSpecificSlot) { targetSlotRef = null; cameraModeContext = "REPORT"; }',
    '  document.getElementById("cameraModal").classList.remove("hidden");',
    '  updateCameraBadgeUI();',
    '  await startCam();',
    '}',
    'function updateCameraBadgeUI() {',
    '  const isBack = currentCameraFacing === "environment";',
    '  document.getElementById("cameraFacingBadge").textContent = isBack ? "กล้องหลัง (Back Camera)" : "กล้องหน้า (Front Camera)";',
    '  const directInput = document.getElementById("directCameraInput");',
    '  if (directInput) directInput.setAttribute("capture", currentCameraFacing);',
    '}',
    'async function startCam() {',
    '  if (cameraStream) { cameraStream.getTracks().forEach(function(t){t.stop();}); cameraStream = null; }',
    '  isLiveCameraReady = false;',
    '  const v = document.getElementById("cameraVideo");',
    '  const fb = document.getElementById("cameraFallbackBox");',
    '  v.classList.add("hidden"); fb.classList.remove("hidden");',
    '  if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) return;',
    '  try {',
    '    try {',
    '      cameraStream = await navigator.mediaDevices.getUserMedia({ video:{ facingMode:{ ideal:currentCameraFacing } }, audio:false });',
    '    } catch(innerErr) {',
    '      cameraStream = await navigator.mediaDevices.getUserMedia({ video:true, audio:false });',
    '    }',
    '    v.srcObject = cameraStream;',
    '    v.setAttribute("playsinline", "true"); v.muted = true;',
    '    await v.play();',
    '    setTimeout(function(){',
    '      if (v.videoWidth > 0 && v.videoHeight > 0) {',
    '        isLiveCameraReady = true;',
    '        v.classList.remove("hidden"); fb.classList.add("hidden");',
    '      }',
    '    }, 250);',
    '  } catch(e) {',
    '    isLiveCameraReady = false;',
    '    v.classList.add("hidden"); fb.classList.remove("hidden");',
    '  }',
    '}',
    'function switchCameraFacing() {',
    '  currentCameraFacing = currentCameraFacing === "environment" ? "user" : "environment";',
    '  updateCameraBadgeUI();',
    '  startCam();',
    '  showToast("สลับเป็น" + (currentCameraFacing === "environment" ? "กล้องหลัง" : "กล้องหน้า") + "แล้ว");',
    '}',
    'function triggerDirectCameraInput() {',
    '  const inp = document.getElementById("directCameraInput");',
    '  inp.value = ""; inp.setAttribute("capture", currentCameraFacing); inp.click();',
    '}',
    'function capturePhotoFromVideo() {',
    '  const v = document.getElementById("cameraVideo");',
    '  if (!isLiveCameraReady || !cameraStream || !v.videoWidth || v.readyState < 2) {',
    '    triggerDirectCameraInput();',
    '    return;',
    '  }',
    '  const c = document.createElement("canvas"); c.width = 800; c.height = 800; const sz = Math.min(v.videoWidth, v.videoHeight);',
    '  const ctx = c.getContext("2d");',
    '  ctx.drawImage(v, (v.videoWidth-sz)/2, (v.videoHeight-sz)/2, sz, sz, 0, 0, 800, 800);',
    '  const dataUrl = c.toDataURL("image/jpeg", 0.88);',
    '  savePhotoFromCamera(dataUrl);',
    '}',
    'function savePhotoFromCamera(dataUrl) {',
    '  if (cameraModeContext === "EQ_CHECKOUT") {',
    '    const dt = document.getElementById("eqCheckoutDate").value || new Date().toISOString().split("T")[0];',
    '    const person = document.getElementById("eqRequester").value || "ช่างผู้เบิก";',
    '    stampDateWatermarkOnPhoto(dataUrl, "เบิกจ่าย/ขึ้นของ", dt, person, function(stamped){',
    '      equipState.checkoutPhoto = stamped; renderEquipMainPhotosUI(); closeCameraModal(); showToast("ถ่ายรูปหลักฐานวันเบิกจ่ายสำเร็จ!");',
    '    });',
    '    return;',
    '  }',
    '  if (cameraModeContext === "EQ_RETURN") {',
    '    const dt = document.getElementById("eqReturnDate").value || new Date().toISOString().split("T")[0];',
    '    const person = document.getElementById("eqReturner").value || document.getElementById("eqRequester").value || "ผู้นำของมาคืน";',
    '    stampDateWatermarkOnPhoto(dataUrl, "นำของมาคืน", dt, person, function(stamped){',
    '      equipState.returnPhoto = stamped; renderEquipMainPhotosUI(); closeCameraModal(); showToast("ถ่ายรูปหลักฐานวันนำของมาคืนสำเร็จ!");',
    '    });',
    '    return;',
    '  }',
    '  if (cameraModeContext === "EQ_ITEM") {',
    '    if (equipState.currentItems[cameraEquipItemIdx]) {',
    '      equipState.currentItems[cameraEquipItemIdx].photo = dataUrl;',
    '      renderEquipItemsTable(); closeCameraModal(); showToast("ถ่ายรูปอุปกรณ์ซ่อมบำรุงเรียบร้อยแล้ว!");',
    '    }',
    '    return;',
    '  }',
    '  if (targetSlotRef) {',
    '    appState.pages[targetSlotRef.pageIndex].slots[targetSlotRef.slotIndex].imageSrc = dataUrl;',
    '    targetSlotRef = null;',
    '  } else {',
    '    placePhotoNext(dataUrl);',
    '  }',
    '  const prev = document.getElementById("lastCapturedPreview");',
    '  if (prev) { prev.src = dataUrl; prev.classList.remove("hidden"); }',
    '  renderAll(); showToast("ถ่ายรูปลงตาราง A4 สำเร็จ!");',
    '}',
    'function handleNativeCameraFallback(e) {',
    '  const f = e.target.files[0]; if (!f) return;',
    '  processImageFileToDataURL(f, function(u){ savePhotoFromCamera(u); });',
    '}',
    'function closeCameraModal() { if (cameraStream) { cameraStream.getTracks().forEach(function(t){t.stop();}); cameraStream = null; } isLiveCameraReady = false; document.getElementById("cameraModal").classList.add("hidden"); }',
    'function openSlotEditModal(pIdx, sIdx) {',
    '  targetSlotRef = { pageIndex:pIdx, slotIndex:sIdx }; const s = appState.pages[pIdx].slots[sIdx];',
    '  document.getElementById("modalSlotStatus").value = s.statusTag; document.getElementById("modalSlotTitle").value = s.subTitle;',
    '  document.getElementById("modalRemarksList").innerHTML = ""; (s.remarks || [""]).forEach(function(r){ addModalRemarkLine(r); });',
    '  document.getElementById("slotEditModal").classList.remove("hidden");',
    '}',
    'function addModalRemarkLine(txt) {',
    '  const d = document.createElement("div"); d.className = "flex gap-1";',
    '  d.innerHTML = "<input type=\\"text\\" value=\\"" + escapeHtml(txt) + "\\" class=\\"modal-rmk flex-1 bg-slate-800 border border-slate-600 rounded px-2 py-1 text-xs text-white\\"><button onclick=\\"this.parentElement.remove()\\" class=\\"px-2 bg-rose-600 text-white rounded text-xs\\">ลบ</button>";',
    '  document.getElementById("modalRemarksList").appendChild(d);',
    '}',
    'function saveSlotEditModal() {',
    '  const s = appState.pages[targetSlotRef.pageIndex].slots[targetSlotRef.slotIndex];',
    '  s.statusTag = document.getElementById("modalSlotStatus").value; s.subTitle = document.getElementById("modalSlotTitle").value;',
    '  const arr = Array.from(document.querySelectorAll(".modal-rmk")).map(function(i){return i.value.trim();}).filter(Boolean);',
    '  s.remarks = arr.length ? arr : ["ดำเนินการเรียบร้อย"]; closeSlotEditModal(); renderA4Pages();',
    '}',
    'function closeSlotEditModal() { document.getElementById("slotEditModal").classList.add("hidden"); }',
    'function toggleJobStatus() {',
    '  appState.jobStatus = appState.jobStatus === "OPEN" ? "CLOSED" : "OPEN";',
    '  document.getElementById("jobStatusBadge").textContent = appState.jobStatus === "CLOSED" ? "ปิดงานแล้ว (CLOSED)" : "เปิดงาน (OPEN)";',
    '  if (appState.jobStatus === "CLOSED") saveToGoogleSheetsDirect();',
    '  renderA4Pages();',
    '}',
    'function buildPayload() {',
    '  syncJobData(); let photos = 0; appState.pages.forEach(function(p){ p.slots.forEach(function(s){ if (s.imageSrc) photos++; }); });',
    '  return { id:"REC-"+Date.now().toString().slice(-5), date:appState.serviceDate, machineCode:appState.machineCode, machineName:appState.machineName, techName:appState.techName, status:appState.jobStatus, alarmCode:appState.alarmCode, comment:appState.comment, photoCount:photos, snapshot:appState };',
    '}',
    'function saveToGoogleSheetsDirect() {',
    '  appState.jobStatus = "รอตรวจสอบ (PENDING)";',
    '  document.getElementById("jobStatusBadge").textContent = "รอตรวจสอบ (PENDING)";',
    '  document.getElementById("jobStatusBadge").className = "px-2 py-0.5 text-[11px] font-semibold rounded-full bg-amber-500/20 text-amber-300 border border-amber-500/40";',
    '  renderA4Pages();',
    '  const payload = buildPayload(); showToast("บันทึกใบงานแล้ว เปลี่ยนสถานะเป็น รอตรวจสอบ!");',
    '  appState.savedSheetsRecords.unshift({ id:payload.id, date:payload.date, machineCode:payload.machineCode, status:appState.jobStatus, temps:appState.tempValues.join(" / ")+" °C", vibs:appState.vibValues.join(" / ")+" mm/s", alarmCode:payload.alarmCode, comment:payload.comment, snapshot:JSON.parse(JSON.stringify(appState)) });',
    '  renderSheetsTable();',
    '  if (typeof google !== "undefined" && google.script && google.script.run) {',
    '    google.script.run.withSuccessHandler(function(res){',
    '      if (res && res.records) appState.savedSheetsRecords = res.records;',
    '      if (res && res.sheetUrl) { const a = document.getElementById("openDirectSheetLink"); a.href = res.sheetUrl; a.classList.remove("hidden"); }',
    '      renderSheetsTable(); showToast("บันทึกข้อมูลสำเร็จ! สถานะ: รอตรวจสอบ (" + res.recordId + ")");',
    '    }).saveReportFromClient(payload);',
    '  }',
    '}',
    'function loadHistoryFromServer() {',
    '  if (typeof google !== "undefined" && google.script && google.script.run) {',
    '    google.script.run.withSuccessHandler(function(res){',
    '      if (res && res.records) { appState.savedSheetsRecords = res.records; renderSheetsTable(); }',
    '      if (res && res.sheetUrl) { const a = document.getElementById("openDirectSheetLink"); a.href = res.sheetUrl; a.classList.remove("hidden"); }',
    '      if (res && res.webAppUrl) { serverWebAppUrl = res.webAppUrl; const inp = document.getElementById("installWebAppUrlInput"); if (inp) inp.value = res.webAppUrl; updateQrCodeFromInput(res.webAppUrl); }',
    '    }).getRecordsFromClient();',
    '  }',
    '}',
    'function openSheetsModal() { document.getElementById("sheetsModal").classList.remove("hidden"); loadHistoryFromServer(); }',
    'function closeSheetsModal() { document.getElementById("sheetsModal").classList.add("hidden"); }',
    'function renderSheetsTable() {',
    '  const tb = document.getElementById("sheetsTableBody");',
    '  if (!appState.savedSheetsRecords.length) { tb.innerHTML = "<tr><td colspan=\\"7\\" class=\\"p-4 text-center text-slate-400\\">ยังไม่มีประวัติในตาราง กดปุ่มบันทึกเพื่อเพิ่มรายการแรก</td></tr>"; return; }',
    '  tb.innerHTML = appState.savedSheetsRecords.map(function(r){',
    '    return "<tr class=\\"hover:bg-slate-800\\"><td class=\\"p-2\\">" + escapeHtml(r.date) + "</td><td class=\\"p-2 font-bold text-cyan-300\\">" + escapeHtml(r.machineCode) + "</td><td class=\\"p-2\\">" + escapeHtml(r.status) + "</td><td class=\\"p-2 text-orange-300\\">" + escapeHtml(r.temps) + "</td><td class=\\"p-2 text-sky-300\\">" + escapeHtml(r.vibs) + "</td><td class=\\"p-2\\">[" + escapeHtml(r.alarmCode) + "] " + escapeHtml(r.comment) + "</td><td class=\\"p-2 text-right\\"><button onclick=\\"loadSnap(\'" + r.id + "\')\\" class=\\"px-2 py-1 rounded bg-cyan-600 text-white\\">เปิดดู</button></td></tr>";',
    '  }).join("");',
    '}',
    'function loadSnap(id) {',
    '  const rec = appState.savedSheetsRecords.find(function(r){return r.id===id;}); if (!rec || !rec.snapshot) return;',
    '  const s = rec.snapshot; appState.machineCode = s.machineCode; appState.serviceDate = s.serviceDate; appState.machineName = s.machineName; appState.techName = s.techName; appState.tempValues = s.tempValues; appState.vibValues = s.vibValues; appState.alarmCode = s.alarmCode; appState.comment = s.comment; appState.pages = s.pages; appState.activePageIndex = 0;',
    '  document.getElementById("inputMachineCode").value = appState.machineCode; document.getElementById("inputServiceDate").value = appState.serviceDate; closeSheetsModal(); renderAll(); showToast("โหลดใบงาน " + appState.machineCode + " สำเร็จ");',
    '}',
    'function exportToPDF() { html2pdf().set({ margin:0, filename:"Report_"+appState.machineCode+".pdf", image:{type:"jpeg",quality:0.95}, html2canvas:{scale:2}, jsPDF:{unit:"mm",format:"a4",orientation:appState.orientation} }).from(document.getElementById("a4PagesContainer")).save(); }',
    'function exportToWord() {',
    '  const blob = new Blob(["\\ufeff<html><body>" + document.getElementById("a4PagesContainer").innerHTML + "</body></html>"], { type:"application/msword" });',
    '  const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = "Report_" + appState.machineCode + ".doc"; a.click();',
    '}',
    'function escapeHtml(s) { return String(s || "").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"); }',
    'function showToast(m) { const b = document.getElementById("toastBox"); document.getElementById("toastMessage").textContent = m; b.classList.remove("hidden"); setTimeout(function(){b.classList.add("hidden");}, 3200); }',
    '<\/script>',
    '</body>',
    '</html>'
  ].join('\n');
}
