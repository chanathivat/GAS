// ===============================================================
//               GLOBAL CONFIGURATION VARIABLES
// ===============================================================
const SPREADSHEET_ID = "1uu86ilkmmwgqKkcY6MhfphGnlo1UbZ2xHW4MeS_FVZw";
const DRIVE_FOLDER_ID = "11MO0ujGCsf9e2P1lJWluiczpWDLDFrGV";
const ADMIN_PASSWORD = "0652370343";

const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

// ===============================================================
//                     WEB APP SERVING
// ===============================================================
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle("CHANATHIVAT Student System")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ===============================================================
//                     DATA FETCHING FUNCTIONS
// ===============================================================
function getSheetNames() {
  try {
    const sheets = ss.getSheets();
    return sheets.map(sheet => sheet.getName());
  } catch (e) {
    return { error: `Error fetching sheet names: ${e.message}` };
  }
}

function getHeaders(sheet) {
  const range = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  return range.getValues()[0];
}

function getStudentData(sheetName) {
  try {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { error: "Sheet not found" };

    const data = sheet.getDataRange().getValues();
    const headers = data.shift(); // Remove header row

    const colMap = {
      no: headers.indexOf("No"),
      id: headers.indexOf("ID"),
      name: headers.indexOf("Name"),
      nickname: headers.indexOf("Nickname"),
      image: headers.indexOf("Image"),
      password: headers.indexOf("Password")
    };

    for (const key in colMap) {
      if (colMap[key] === -1) {
        return { error: `Column "${key.charAt(0).toUpperCase() + key.slice(1)}" not found in sheet "${sheetName}".` };
      }
    }

    const students = data.map(row => ({
      no: row[colMap.no],
      id: row[colMap.id],
      name: row[colMap.name],
      nickname: row[colMap.nickname],
      image: row[colMap.image],
    }));

    return { students: students };
  } catch (e) {
    return { error: `Error fetching student data: ${e.message}` };
  }
}

function verifyStudentPassword(sheetName, studentId, password) {
  try {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { success: false, message: "Sheet not found" };

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idCol = headers.indexOf("ID");
    const passwordCol = headers.indexOf("Password");

    if (idCol === -1 || passwordCol === -1) {
      return { success: false, message: "ID or Password column not found." };
    }

    const studentRow = data.find(row => row[idCol] == studentId);

    if (!studentRow) {
      return { success: false, message: "Student not found." };
    }
    
    // *** START: CODE FIX FOR ADMIN EDIT ***
    // This allows an admin (who is already authenticated on the client-side)
    // to fetch student data for editing without needing the student's actual password.
    if (password === 'dummy_password_for_data_fetch') {
        const studentData = {};
        headers.forEach((header, index) => {
            studentData[header] = studentRow[index];
        });
        return { success: true, studentData: studentData, headers: headers };
    }
    // *** END: CODE FIX FOR ADMIN EDIT ***

    if (studentRow[passwordCol] == password) {
      const studentData = {};
      headers.forEach((header, index) => {
        studentData[header] = studentRow[index];
      });
      return { success: true, studentData: studentData, headers: headers };
    } else {
      return { success: false, message: "Incorrect password." };
    }
  } catch (e) {
    return { success: false, message: `Error: ${e.message}` };
  }
}

// ===============================================================
//                     DATA MODIFICATION FUNCTIONS
// ===============================================================

function updateStudentPassword(sheetName, studentId, newPassword) {
  try {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { success: false, message: "Sheet not found" };

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idCol = headers.indexOf("ID");
    const passwordCol = headers.indexOf("Password");
    
    if (idCol === -1 || passwordCol === -1) return { success: false, message: "Column 'ID' or 'Password' not found." };

    for (let i = 1; i < data.length; i++) {
      if (data[i][idCol] == studentId) {
        sheet.getRange(i + 1, passwordCol + 1).setValue(newPassword);
        return { success: true, message: "Password updated successfully." };
      }
    }
    return { success: false, message: "Student ID not found." };
  } catch (e) {
    return { success: false, message: `Error: ${e.message}` };
  }
}

function uploadImage(fileData, oldImageId) {
  try {
    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);

    if (oldImageId) {
      try {
        const oldFile = DriveApp.getFileById(oldImageId);
        oldFile.setTrashed(true);
      } catch (e) {
        console.log(`Could not delete old file with ID ${oldImageId}. It might have been deleted already. Error: ${e.message}`);
      }
    }

    const decoded = Utilities.base64Decode(fileData.base64);
    const blob = Utilities.newBlob(decoded, fileData.mimeType, fileData.name);
    const newFile = folder.createFile(blob);
    newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return { success: true, fileId: newFile.getId() };
  } catch (e) {
    return { success: false, message: `Image upload failed: ${e.message}` };
  }
}

function updateStudentData(sheetName, studentId, updatedData) {
   try {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { success: false, message: "Sheet not found" };
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idCol = headers.indexOf("ID");
    
    if (idCol === -1) return { success: false, message: "Column 'ID' not found." };
    
    let studentRowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][idCol] == studentId) {
        studentRowIndex = i + 1;
        break;
      }
    }
    
    if (studentRowIndex === -1) return { success: false, message: "Student ID not found." };

    headers.forEach((header, index) => {
      if (updatedData.hasOwnProperty(header)) {
        sheet.getRange(studentRowIndex, index + 1).setValue(updatedData[header]);
      }
    });

    return { success: true, message: "Data updated successfully." };
  } catch (e) {
    return { success: false, message: `Error updating data: ${e.message}` };
  }
}

function addNewScoreColumn(sheetName, topicName) {
  try {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { success: false, message: "Sheet not found" };

    const lastCol = sheet.getLastColumn();
    sheet.getRange(1, lastCol + 1).setValue(topicName);
    return { success: true, message: `Column '${topicName}' added successfully.`};
  } catch (e) {
    return { success: false, message: `Error adding column: ${e.message}` };
  }
}

// ===============================================================
//                     ADMIN VERIFICATION
// ===============================================================
function verifyAdminPassword(password) {
  return { isAdmin: password === ADMIN_PASSWORD };
}
