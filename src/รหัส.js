/**
 * Smart School Achievement & Portfolio System
 * Senior Developer Edition - Full Version with Academic Year & Robust Save
 */

// --- CONFIGURATION ---
const APP_CONFIG = {
  DB_NAME: "SMART_SCHOOL_DB_v1", 
  FOLDER_NAME: "School_Achievement_Images", 
  ADMIN_LIST: ["admin"], 
};

// --- CORE FUNCTIONS (SPA) ---

function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Smart School Achievement System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// --- DATABASE & INITIALIZATION ---

function getDbSpreadsheet() {
  const files = DriveApp.getFilesByName(APP_CONFIG.DB_NAME);
  if (files.hasNext()) {
    return SpreadsheetApp.open(files.next());
  } else {
    const ss = SpreadsheetApp.create(APP_CONFIG.DB_NAME);
    initializeDatabase(ss);
    return ss;
  }
}

function initializeDatabase(ss = null) {
  if (!ss) ss = getDbSpreadsheet();

  // 1. Config Sheet
  let sheetConfig = ss.getSheetByName('Config');
  if (!sheetConfig) {
    sheetConfig = ss.insertSheet('Config');
    sheetConfig.appendRow(['Key', 'Value']);
    sheetConfig.appendRow(['SchoolName', 'โรงเรียนอนุบาลเมืองอาจสามารถ']);
    sheetConfig.appendRow(['SchoolYear', '2569']);
    sheetConfig.appendRow(['CurrentTerm', '1']);
  }

  // 2. Users Sheet
  let sheetUsers = ss.getSheetByName('Users');
  if (!sheetUsers) {
    sheetUsers = ss.insertSheet('Users');
    sheetUsers.appendRow(['Username', 'Password', 'Name', 'Role', 'ImageURL']);
    sheetUsers.appendRow(['admin', '1234', 'Admin Master', 'admin', 'https://via.placeholder.com/150']);
    sheetUsers.appendRow(['teacher1', '1234', 'คุณครูตัวอย่าง', 'teacher', 'https://via.placeholder.com/150']);
  }

  // 3. Achievements Sheet
  let sheetAch = ss.getSheetByName('Achievements');
  if (!sheetAch) {
    sheetAch = ss.insertSheet('Achievements');
    // Header (Column 1-11)
    sheetAch.appendRow([
      'ID', 
      'Title', 
      'Category', 
      'Level',    
      'OwnerName', 
      'Date', 
      'ImageID', 
      'Description',
      'CreatedBy', 
      'Timestamp',
      'AcademicYear'
    ]);
  }
}

function createResponse(status, data = null, message = "") {
  return { status: status, data: data, message: message };
}

// --- CONFIGURATION SERVICE ---

function getSystemConfig() {
  try {
    const ss = getDbSpreadsheet();
    const sheet = ss.getSheetByName('Config');
    const data = sheet.getDataRange().getValues();
    const config = {};
    for (let i = 1; i < data.length; i++) {
      config[data[i][0]] = data[i][1];
    }
    return createResponse("success", config);
  } catch (e) {
    return createResponse("error", null, e.toString());
  }
}

function saveSystemConfig(configData) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return createResponse("error", null, "Server busy");
  try {
    const ss = getDbSpreadsheet();
    const sheet = ss.getSheetByName('Config');
    const data = sheet.getDataRange().getValues();
    const updateKey = (key, val) => {
      let found = false;
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] == key) {
          sheet.getRange(i + 1, 2).setValue(val);
          found = true;
          break;
        }
      }
      if (!found) sheet.appendRow([key, val]);
    };
    if (configData.SchoolYear) updateKey('SchoolYear', configData.SchoolYear);
    if (configData.CurrentTerm) updateKey('CurrentTerm', configData.CurrentTerm);
    return createResponse("success", null, "บันทึกการตั้งค่าเรียบร้อย");
  } catch (e) {
    return createResponse("error", null, e.toString());
  } finally {
    lock.releaseLock();
  }
}

// --- ACHIEVEMENT SERVICE (UPDATED LOGIC) ---

function handleAchievementSave(payload) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return createResponse("error", null, "ระบบกำลังทำงานหนัก กรุณาลองใหม่");

  try {
    const ss = getDbSpreadsheet();
    const sheet = ss.getSheetByName('Achievements');
    const timestamp = new Date();
    
    let imageId = payload.currentImageId || "";
    if (payload.imageFile && payload.imageFile.data) {
      imageId = uploadImageToDrive(payload.imageFile.data, payload.imageFile.name, payload.imageFile.mimeType);
    }

    // กำหนด ID: ถ้ามีส่งมา (แก้ไข) ให้ใช้ตัวเดิม, ถ้าไม่มี (เพิ่มใหม่) ให้สร้างใหม่
    const recordId = payload.id ? String(payload.id) : Utilities.getUuid();

    const rowData = [
      recordId,
      payload.title,
      payload.category,
      payload.level,
      payload.ownerName,
      payload.date, 
      imageId,
      payload.description || "",
      payload.createdBy,
      timestamp,
      payload.academicYear || ""
    ];

    // ตรวจสอบ Mode: แก้ไข หรือ เพิ่มใหม่
    if (payload.id) {
      // --- MODE: EDIT ---
      const data = sheet.getDataRange().getValues();
      let rowIndex = -1;
      
      // วนลูปหา ID ที่ตรงกัน (เริ่มที่แถว 2 เพราะแถว 1 คือ Header)
      for (let i = 1; i < data.length; i++) {
        // ใช้ String() ครอบเพื่อความชัวร์ในการเปรียบเทียบ
        if (String(data[i][0]) === String(payload.id)) {
          rowIndex = i + 1; // Index ใน Sheet เริ่มที่ 1
          break;
        }
      }

      if (rowIndex > 0) {
        // เจอแถวเดิม -> อัปเดตข้อมูลทับ
        const range = sheet.getRange(rowIndex, 1, 1, rowData.length);
        range.setValues([rowData]);
      } else {
        // มี ID ส่งมา แต่หาไม่เจอ (อาจถูกลบไปแล้ว) -> เพิ่มเป็นแถวใหม่ (Upsert)
        sheet.appendRow(rowData);
      }
    } else {
      // --- MODE: ADD ---
      // ไม่มี ID ส่งมา -> เพิ่มแถวใหม่ทันที
      sheet.appendRow(rowData);
    }

    return createResponse("success", null, "บันทึกข้อมูลสำเร็จ");

  } catch (e) {
    return createResponse("error", null, "เกิดข้อผิดพลาด: " + e.toString());
  } finally {
    lock.releaseLock();
  }
}

function getPublicAchievements() {
  try {
    const ss = getDbSpreadsheet();
    const sheet = ss.getSheetByName('Achievements');
    const data = sheet.getDataRange().getValues();
    const results = [];

    for (let i = 1; i < data.length; i++) {
      let row = data[i];
      results.push({
        id: row[0],
        title: row[1],
        category: row[2],
        level: row[3],
        ownerName: row[4],
        date: row[5] instanceof Date ? Utilities.formatDate(row[5], Session.getScriptTimeZone(), "yyyy-MM-dd") : row[5],
        imageId: row[6],
        description: row[7],
        createdBy: row[8],
        academicYear: row[10] || ""
      });
    }
    results.sort((a, b) => new Date(b.date) - new Date(a.date));
    return createResponse("success", results, "ดึงข้อมูลสำเร็จ");
  } catch (e) {
    return createResponse("error", null, e.toString());
  }
}

function deleteAchievement(id) {
   const lock = LockService.getScriptLock();
   if (!lock.tryLock(5000)) return createResponse("error", null, "Server busy");
   try {
    const ss = getDbSpreadsheet();
    const sheet = ss.getSheetByName('Achievements');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) { // ใช้ String compare เพื่อความชัวร์
        sheet.deleteRow(i + 1);
        return createResponse("success", null, "ลบข้อมูลเรียบร้อย");
      }
    }
    return createResponse("error", null, "ไม่พบข้อมูล");
   } catch(e) {
     return createResponse("error", null, e.toString());
   } finally {
     lock.releaseLock();
   }
}

// --- USER MANAGEMENT & AUTH SERVICE ---

function getAllUsers() {
  try {
    const ss = getDbSpreadsheet();
    const sheet = ss.getSheetByName('Users');
    const data = sheet.getDataRange().getValues();
    const users = [];
    for (let i = 1; i < data.length; i++) {
      users.push({
        username: data[i][0],
        password: data[i][1],
        name: data[i][2],
        role: data[i][3],
        image: data[i][4]
      });
    }
    return createResponse("success", users);
  } catch (e) {
    return createResponse("error", null, e.toString());
  }
}

function saveUser(data) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return createResponse("error", null, "Server busy");
  try {
    const ss = getDbSpreadsheet();
    const sheet = ss.getSheetByName('Users');
    const rows = sheet.getDataRange().getValues();
    let rowIndex = -1;
    
    // Check Duplicate / Find Row
    if (!data.isEdit) {
      for(let i=1; i<rows.length; i++) {
        if(String(rows[i][0]) === String(data.username)) return createResponse("error", null, "ชื่อผู้ใช้นี้มีอยู่แล้ว");
      }
    } else {
      for(let i=1; i<rows.length; i++) {
        if(String(rows[i][0]) === String(data.username)) {
          rowIndex = i + 1;
          break;
        }
      }
      if (rowIndex == -1) return createResponse("error", null, "ไม่พบผู้ใช้");
    }

    let imageUrl = data.image || "https://via.placeholder.com/150";
    if (data.imageFile && data.imageFile.data) {
       let imgId = uploadImageToDrive(data.imageFile.data, "user_" + data.username, data.imageFile.mimeType);
       imageUrl = `https://lh3.googleusercontent.com/d/${imgId}=s200`;
    }

    if (data.isEdit) {
      if(data.password) sheet.getRange(rowIndex, 2).setValue(data.password);
      sheet.getRange(rowIndex, 3).setValue(data.name);
      sheet.getRange(rowIndex, 4).setValue(data.role);
      sheet.getRange(rowIndex, 5).setValue(imageUrl);
    } else {
      sheet.appendRow([data.username, data.password, data.name, data.role, imageUrl]);
    }
    return createResponse("success", null, "บันทึกสำเร็จ");
  } catch(e) {
    return createResponse("error", null, e.toString());
  } finally {
    lock.releaseLock();
  }
}

function deleteUser(username) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return createResponse("error", null, "Server busy");
  try {
    const ss = getDbSpreadsheet();
    const sheet = ss.getSheetByName('Users');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(username)) {
        sheet.deleteRow(i + 1);
        return createResponse("success", null, "ลบสำเร็จ");
      }
    }
    return createResponse("error", null, "ไม่พบผู้ใช้งาน");
  } catch(e) {
    return createResponse("error", null, e.toString());
  } finally {
    lock.releaseLock();
  }
}

function updateUserProfile(username, newName, newImageBase64) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return createResponse("error", null, "System busy");
  try {
    const ss = getDbSpreadsheet();
    const sheet = ss.getSheetByName('Users');
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    let currentImage = "";
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(username)) {
        rowIndex = i + 1;
        currentImage = data[i][4];
        break;
      }
    }
    if (rowIndex == -1) return createResponse("error", null, "User not found");
    let newImageUrl = currentImage;
    if (newImageBase64 && newImageBase64.includes('base64')) {
       let imgId = uploadImageToDrive(newImageBase64, `profile_${username}_${Date.now()}.png`, "image/png");
       newImageUrl = `https://lh3.googleusercontent.com/d/${imgId}=s200`;
    }
    sheet.getRange(rowIndex, 3).setValue(newName);
    sheet.getRange(rowIndex, 5).setValue(newImageUrl);
    return createResponse("success", { name: newName, image: newImageUrl }, "Success");
  } catch (e) {
    return createResponse("error", null, "Error: " + e.toString());
  } finally {
    lock.releaseLock();
  }
}

function changeUserPassword(username, oldPass, newPass) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return createResponse("error", null, "System busy");
  try {
    const ss = getDbSpreadsheet();
    const sheet = ss.getSheetByName('Users');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(username)) {
        if (String(data[i][1]) === String(oldPass)) {
           sheet.getRange(i + 1, 2).setValue(newPass);
           return createResponse("success", null, "Password changed");
        } else {
           return createResponse("error", null, "Incorrect old password");
        }
      }
    }
    return createResponse("error", null, "User not found");
  } catch (e) {
    return createResponse("error", null, e.toString());
  } finally {
    lock.releaseLock();
  }
}

function loginUser(username, password) {
  const ss = getDbSpreadsheet();
  const sheet = ss.getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  for(let i=1; i<data.length; i++) {
    if(String(data[i][0]) === String(username) && String(data[i][1]) === String(password)) {
      return createResponse("success", {
        username: data[i][0],
        name: data[i][2],
        role: data[i][3],
        image: data[i][4]
      }, "Login success");
    }
  }
  return createResponse("error", null, "Invalid credentials");
}

function uploadImageToDrive(base64Data, fileName, mimeType) {
  try {
    const folders = DriveApp.getFoldersByName(APP_CONFIG.FOLDER_NAME);
    let folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(APP_CONFIG.FOLDER_NAME);
    folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const splitBase64 = base64Data.split(',');
    const data = splitBase64.length > 1 ? splitBase64[1] : splitBase64[0];
    const decoded = Utilities.base64Decode(data);
    const blob = Utilities.newBlob(decoded, mimeType, fileName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getId();
  } catch (e) {
    throw new Error("Upload Failed: " + e.message);
  }
}