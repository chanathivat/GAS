// ===== ค่าคงที่ =====
// ID ของ Google Sheet ที่ใช้เป็นฐานข้อมูล
const SPREADSHEET_ID = '1uu86ilkmmwgqKkcY6MhfphGnlo1UbZ2xHW4MeS_FVZw'; 
// ID ของโฟลเดอร์ใน Google Drive ที่ใช้เก็บรูปนักเรียน
const DRIVE_FOLDER_ID = '11MO0ujGCsf9e2P1lJWluiczpWDLDFrGV'; 

// เปิดการเข้าถึง Spreadsheet และ Drive Folder
const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
const driveFolder = DriveApp.getFolderById(DRIVE_FOLDER_ID);

// ===== จุดเริ่มต้นของ Web App (Entry Points) =====



function doGet(e) {
  try {
    if (!e.parameter.action) {
      return HtmlService.createHtmlOutputFromFile('index')
        .setTitle('CHANATHIVAT Student Information System')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    const action = e.parameter.action;
    let result;

    switch (action) {
      case 'getSheets':
        result = getSheetNames();
        break;
      case 'getStudents':
        const sheetName = e.parameter.sheet;
        if (!sheetName) throw new Error("กรุณาระบุชื่อชีท");
        result = getStudents(sheetName);
        break;
      default:
        throw new Error("Action ไม่ถูกต้อง");
    }

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, ...result }))
      .setMimeType(ContentService.MimeType.JSON)
      .setHeader('Access-Control-Allow-Origin', '*');

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: error.message }))
      .setMimeType(ContentService.MimeType.JSON)
      .setHeader('Access-Control-Allow-Origin', '*');
  }
}

function doPost(e) {
  try {
    const action = e.parameter.action || JSON.parse(e.postData.contents).action;
    let data;
    let result;

    if (action !== 'uploadImage') {
      data = JSON.parse(e.postData.contents);
    }

    switch (action) {
      case 'updateStudent':
        result = updateStudent(data.sheet, data.student);
        break;
      case 'addStudent':
        result = addStudent(data.sheet, data.student);
        break;
      case 'addScoreColumn':
        result = addScoreColumn(data.sheet, data.columnName);
        break;
      case 'changePassword':
        result = changePassword(data.sheet, data.studentId, data.newPassword);
        break;
      case 'uploadImage':
        result = uploadImage(e);
        break;
      default:
        throw new Error("Action ไม่ถูกต้อง");
    }

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, ...result }))
      .setMimeType(ContentService.MimeType.JSON)
      .setHeader('Access-Control-Allow-Origin', '*');

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: error.message }))
      .setMimeType(ContentService.MimeType.JSON)
      .setHeader('Access-Control-Allow-Origin', '*');
  }
}

// === Helper Functions ===
// ตัวอย่างฟังก์ชัน getSheetNames() และอื่น ๆ ตามที่คุณมีอยู่ในโค้ดของคุณ

function getSheetNames() {
  const sheets = ss.getSheets();
  const sheetNames = sheets.map(sheet => sheet.getName());
  return { sheets: sheetNames };
}

function getStudents(sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`ไม่พบชีทชื่อ: ${sheetName}`);
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  const students = data.map(row => {
    const studentObj = {};
    headers.forEach((header, i) => {
      studentObj[header] = row[i];
    });
    return studentObj;
  });
  return { students: students, headers: headers };
}

// (ฟังก์ชันอื่น ๆ เช่น updateStudent, addStudent, changePassword ... ให้ใช้ของคุณเดิม)




// ===== ฟังก์ชันจัดการข้อมูล (Helper Functions) =====

/**
 * ดึงรายชื่อชีททั้งหมดใน Spreadsheet
 */
function getSheetNames() {
  const sheets = ss.getSheets();
  const sheetNames = sheets.map(sheet => sheet.getName());
  return { sheets: sheetNames };
}

/**
 * ดึงข้อมูลนักเรียนทั้งหมดจากชีทที่ระบุ
 */
function getStudents(sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`ไม่พบชีทชื่อ: ${sheetName}`);
  
  const data = sheet.getDataRange().getValues();
  const headers = data.shift(); // เอาแถวแรกสุดเป็น headers
  const students = data.map(row => {
    const studentObj = {};
    headers.forEach((header, i) => {
      studentObj[header] = row[i];
    });
    return studentObj;
  });

  return { students: students, headers: headers };
}

/**
 * อัปเดตข้อมูลนักเรียน
 */
function updateStudent(sheetName, studentData) {
  const sheet = ss.getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idColumnIndex = headers.indexOf('ID');
  
  // หาแถวของนักเรียนที่ต้องการอัปเดต (เริ่มนับจาก 1, +1 เพราะ data ไม่มี header)
  const rowIndex = data.findIndex(row => row[idColumnIndex] == studentData.ID) + 1; 

  if (rowIndex > 0) {
    const newRow = headers.map(header => studentData[header] || '');
    sheet.getRange(rowIndex, 1, 1, newRow.length).setValues([newRow]);
    return { message: "อัปเดตข้อมูลสำเร็จ" };
  } else {
    throw new Error("ไม่พบรหัสนักเรียนที่ต้องการอัปเดต");
  }
}

/**
 * เพิ่มนักเรียนใหม่
 */
function addStudent(sheetName, studentData) {
  const sheet = ss.getSheetByName(sheetName);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const newRow = headers.map(header => studentData[header] || '');
  sheet.appendRow(newRow);
  return { message: "เพิ่มนักเรียนสำเร็จ" };
}


/**
 * เพิ่มคอลัมน์คะแนนใหม่
 */
function addScoreColumn(sheetName, columnName) {
  const sheet = ss.getSheetByName(sheetName);
  const lastColumn = sheet.getLastColumn();
  sheet.getRange(1, lastColumn + 1).setValue(columnName);
  return { message: "เพิ่มหัวข้อคะแนนสำเร็จ" };
}

/**
 * เปลี่ยนรหัสผ่านของนักเรียน
 */
function changePassword(sheetName, studentId, newPassword) {
  const sheet = ss.getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idColumnIndex = headers.indexOf('ID');
  const passwordColumnIndex = headers.indexOf('Password');

  const rowIndex = data.findIndex(row => row[idColumnIndex] == studentId) + 1;

  if (rowIndex > 0) {
    // +1 เพราะ getRange เริ่มที่ 1
    sheet.getRange(rowIndex, passwordColumnIndex + 1).setValue(newPassword);
    return { message: "เปลี่ยนรหัสผ่านสำเร็จ" };
  } else {
    throw new Error("ไม่พบรหัสนักเรียน");
  }
}

/**
 * อัปโหลดไฟล์รูปภาพไปที่ Google Drive
 */
function uploadImage(e) {
  const fileBlob = e.postData.contents;
  const blob = Utilities.newBlob(Utilities.base64Decode(fileBlob), e.parameter.mimeType, e.parameter.fileName);
  
  // ลบรูปเก่าถ้ามี
  if (e.parameter.oldImageId) {
    try {
      DriveApp.getFileById(e.parameter.oldImageId).setTrashed(true);
    } catch (err) {
      console.log("ไม่สามารถลบรูปเก่าได้: " + err.message);
    }
  }

  const file = driveFolder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  
  return { imageId: file.getId() };
}
