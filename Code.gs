function doGet(e) {
  let page = e?.parameter?.page || 'Index';
  return HtmlService.createTemplateFromFile(page).evaluate()
    .setTitle('ระบบจัดการสำนักงานออนไลน์')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ==========================================
// 🛠️ ฟังก์ชันเครื่องมือพื้นฐาน (Utilities)
// ==========================================
function formatDateSafe(dateVal) { 
  if (!dateVal) return "-";
  try { 
    return Utilities.formatDate(new Date(dateVal), Session.getScriptTimeZone(), "dd/MM/yyyy"); 
  } catch(e) { 
    return dateVal.toString();
  } 
}

function formatDateForInput(dateVal) { 
  if (!dateVal) return "";
  try { 
    return Utilities.formatDate(new Date(dateVal), Session.getScriptTimeZone(), "yyyy-MM-dd"); 
  } catch(e) { 
    return "";
  } 
}

function formatDateTimeSafe(dateVal) {
  if (!dateVal) return "-"; 
  try { 
    return Utilities.formatDate(new Date(dateVal), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  } catch(e) { 
    return dateVal.toString(); 
  } 
}

function cleanFolderId(input) { 
  if (!input) return "";
  let str = String(input).trim(); 
  
  let match1 = str.match(/folders\/([a-zA-Z0-9-_]+)/); 
  if (match1) return match1[1]; 
  
  let match2 = str.match(/id=([a-zA-Z0-9-_]+)/); 
  if (match2) return match2[1];
  
  return str; 
}

function cleanTgText(text) { 
  if(!text) return "-"; 
  return String(text).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

// ==========================================
// 📊 ระบบประวัติการใช้งาน (Audit Log) & Telegram
// ==========================================
function writeAuditLog(userEmail, action, detail) {
  try {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("AuditLogs");
    if (!sheet) { 
      sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("AuditLogs"); 
      sheet.appendRow(["Timestamp", "User Email", "Action", "Detail"]);
    }
    sheet.appendRow([new Date(), userEmail || "System", action, detail]);
  } catch(e) {
    console.error("AuditLog Error: " + e);
  } 
}

function getAuditLogsAdmin(opt_nocache) {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("AuditLogs");
  if (!sheet) return [];
  
  let data = sheet.getDataRange().getDisplayValues(); 
  let res = [];
  for(let i = data.length - 1; i >= 1; i--) { 
    res.push({ 
      timestamp: data[i][0] || "-", 
      user: data[i][1] || "-", 
      action: data[i][2] || "-", 
      detail: data[i][3] || "-" 
    });
    if(res.length >= 300) break; 
  }
  return res;
}

function getUserPersonalTgChatId(email) {
  try {
    if (!email) return "";
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
    if(!sheet) return "";
    
    const data = sheet.getDataRange().getValues();
    for(let i = 1; i < data.length; i++) { 
      if(data[i][0] === email && data[i][13]) { 
        return String(data[i][13]).trim();
      } 
    }
  } catch(e) {
    console.error("getUserPersonalTgChatId Error: " + e);
  } 
  return "";
}

function sendTelegram(message, attachmentUrl = "", pdfUrl = "", targetChatId = "", systemType = "leave") {
  const settings = getAllSettings();
  let token = "";
  let defaultChatId = "";

  // 1. ตรวจสอบว่ามีการตั้งค่า TELEGRAM แบบแยกระบบ (JSON) อยู่หรือไม่
  if (settings.TELEGRAM) {
    try {
      const tgData = JSON.parse(settings.TELEGRAM);
      if (tgData[systemType]) {
        token = String(tgData[systemType].token || "").trim();
        defaultChatId = String(tgData[systemType].chatId || "").trim();
      }
    } catch (e) {
      console.error("Error parsing TELEGRAM JSON: " + e);
    }
  }

  // 2. ถ้ายังไม่ได้ตั้งค่าแบบใหม่ (หรือเป็นค่าว่าง) และเป็นการแจ้งเตือนของ "ระบบลางาน" ให้ไปดึงจากคีย์เก่า
  if (!token && systemType === "leave") {
    token = String(settings.telegramToken || "").trim();
    defaultChatId = String(settings.telegramChatId || "").trim();
  }

  // กำหนด Chat ID
  const chatId = targetChatId ? targetChatId : defaultChatId;
  
  if (!token || !chatId) return;

  const webAppUrl = ScriptApp.getService().getUrl();
  let keyboard = [[{ text: "💻 เข้าสู่ระบบ", url: webAppUrl }]];

  if (attachmentUrl && attachmentUrl !== "ไม่มีไฟล์แนบ" && String(attachmentUrl).startsWith("http")) { 
    keyboard.push([{ text: "📎 ดูเอกสารแนบ", url: attachmentUrl }]);
  }
  if (pdfUrl && String(pdfUrl).startsWith("http")) { 
    keyboard.push([{ text: "📄 โหลดไฟล์ PDF", url: pdfUrl }]);
  }
  
  const url = `https://api.telegram.org/bot${token}/sendMessage`;
  const payload = { 
    chat_id: chatId, 
    text: message, 
    parse_mode: "HTML", 
    reply_markup: { inline_keyboard: keyboard } 
  };

  const options = { 
    method: "post", 
    contentType: "application/json", 
    payload: JSON.stringify(payload), 
    muteHttpExceptions: true 
  };

  try { 
    UrlFetchApp.fetch(url, options); 
  } catch (e) { 
    console.error("Telegram Error: " + e);
  }
}

function getUserInfoByEmail(email) {
  if (!email) return { name: "ไม่ระบุ", tgChatId: "" };
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  if (!sheet) return { name: email, tgChatId: "" };
  
  const data = sheet.getDataRange().getValues();
  const tgColIndex = 13; // หากแก้คอลัมน์ Telegram อย่าลืมปรับ Index ตรงนี้
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email) {
      return {
        name: data[i][1] || email,
        tgChatId: data[i][tgColIndex] ? data[i][tgColIndex].toString().trim() : ""
      };
    }
  }
  return { name: email, tgChatId: "" };
}

// ==========================================
// 👤 ระบบผู้ใช้งาน & โปรไฟล์ (Users & Profile)
// ==========================================
function authenticateUser(email, password) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
    if (!sheet) return { success: false, message: "ไม่พบฐานข้อมูลผู้ใช้งาน" };
    
    const data = sheet.getDataRange().getDisplayValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === email) {
        
        // รหัสผ่าน Index 16
        const dbPassword = data[i][16] ? data[i][16].toString().trim() : "";
        const inputPassword = password ? password.toString().trim() : "";

        if (dbPassword !== "" && dbPassword !== inputPassword) {
          return { success: false, message: "รหัสผ่านไม่ถูกต้อง กรุณาลองใหม่" };
        }

        return { 
          success: true, 
          profile: { 
            email: data[i][0], 
            fullName: data[i][1], 
            phone: data[i][2], 
            position: data[i][3], 
            affiliation: data[i][4], 
            address: data[i][5], 
            signature: data[i][6], 
            avatar: data[i][7], 
            roles: String(data[i][8] || 'User').split(',').map(r => r.trim()), 
            accumulatedLeave: parseFloat(data[i][9]) || 0, 
            personalTgChatId: data[i][13] || "",
            employmentType: data[i][15] || "อื่นๆ"
          } 
        };
      }
    }
    return { success: false, message: "ไม่พบอีเมลนี้ในระบบ" };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function saveProfile(profileObj) {
  try {
    writeAuditLog(profileObj.email, "Update Profile", "อัปเดตข้อมูลส่วนตัวและลายเซ็น");
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
    if (!sheet) return { success: false, message: "ไม่พบแผ่นงาน Users" };

    const data = sheet.getDataRange().getValues();
    
    let rowIndex = -1; 
    let existingData = Array(17).fill(""); // สร้างอาร์เรย์ว่างเผื่อไว้เพื่อหลีกเลี่ยง undefined
    
    for (let i = 1; i < data.length; i++) { 
      if (data[i][0] === profileObj.email) { 
        rowIndex = i + 1;
        existingData = data[i]; 
        break; 
      } 
    }
    
    let dEmail = existingData[10] !== undefined ? existingData[10] : ""; 
    let dStart = existingData[11] !== undefined ? existingData[11] : ""; 
    let dEnd = existingData[12] !== undefined ? existingData[12] : "";
    
    const rowData = [
      profileObj.email, 
      profileObj.fullName || "", 
      profileObj.phone || "", 
      profileObj.position || "", 
      profileObj.affiliation || "", 
      profileObj.address || "", 
      profileObj.signature || "", 
      profileObj.avatar || "", 
      existingData[8] || "User", 
      profileObj.accumulatedLeave || 0, 
      dEmail, 
      dStart, 
      dEnd, 
      profileObj.personalTgChatId || "",
      existingData[14] || "", // ข้อมูลเดิมเพื่อไม่ให้สูญหาย
      existingData[15] || "",
      existingData[16] || ""
    ];
    
    if (rowIndex > -1) {
      // เขียนทับด้วยความยาวคอลัมน์ที่แน่นอน
      sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }
    
    SpreadsheetApp.flush();
    return { success: true, message: "อัปเดตข้อมูลโปรไฟล์เรียบร้อยแล้ว" };
  } catch (e) { 
    return { success: false, message: e.toString() };
  }
}
// ==========================================
// ⚙️ ระบบตั้งค่า (Settings)
// ==========================================
function getAllSettings(opt_nocache) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
    if (!sheet) { 
      const newSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Settings"); 
      newSheet.appendRow(["Key", "Value"]); 
      return {};
    }
    
    const data = sheet.getDataRange().getValues(); 
    let settings = {};
    for (let i = 1; i < data.length; i++) { 
      if (!data[i][0]) continue; // ข้ามบรรทัดที่ Key ว่าง
      let key = data[i][0].toString().trim();
      let val = data[i][1]; 
      if (val instanceof Date) {
        val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
      }
      settings[key] = val; 
    }
    return settings;
  } catch (e) {
    console.error("getAllSettings Error: " + e);
    return {};
  }
}

// ==========================================
// 📅 ระบบวันหยุด (Holidays)
// ==========================================
function syncThaiHolidays(yearStr, actorEmail) {
  try {
    writeAuditLog(actorEmail, "Sync Holidays", `ดึงข้อมูลวันหยุดอัตโนมัติ ปี ${yearStr}`);
    const year = parseInt(yearStr);
    if (isNaN(year)) return { success: false, message: "ระบุปีไม่ถูกต้อง" };
    
    const calIds = [
      'th.th#holiday@group.v.calendar.google.com', 
      'th.thai#holiday@group.v.calendar.google.com', 
      'en.th#holiday@group.v.calendar.google.com'
    ];
    let cal = null;
    for (let id of calIds) { 
      cal = CalendarApp.getCalendarById(id);
      if (!cal) { 
        try { cal = CalendarApp.subscribeToCalendar(id); } catch(e) {} 
      } 
      if (cal) break;
    }
    
    if (!cal) return { success: false, message: "ไม่สามารถเชื่อมต่อปฏิทินได้ กรุณาตั้งค่าแบบ Manual" };
    
    const startDate = new Date(year, 0, 1); 
    const endDate = new Date(year, 11, 31); 
    const events = cal.getEvents(startDate, endDate);
    
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Holidays");
    if (!sheet) { 
      sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Holidays"); 
      sheet.appendRow(["Date", "Name"]);
    }
    
    const existingData = sheet.getDataRange().getValues(); 
    let existingDates = [];
    for(let i = 1; i < existingData.length; i++) { 
      if(existingData[i][0]) { 
        existingDates.push(formatDateForInput(existingData[i][0]));
      } 
    }

    let addedCount = 0;
    events.forEach(e => { 
      let dateStr = formatDateForInput(e.getStartTime()); 
      if (!existingDates.includes(dateStr)) { 
        sheet.appendRow([e.getStartTime(), e.getTitle()]); 
        existingDates.push(dateStr); 
        addedCount++; 
      } 
    });
    SpreadsheetApp.flush(); 
    return { success: true, message: `ดึงข้อมูลสำเร็จ! เพิ่มวันหยุดใหม่ ${addedCount} วัน` };
  } catch (error) { 
    return { success: false, message: error.toString() };
  }
}

function getHolidaysAdmin(opt_nocache) {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Holidays");
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues(); 
  let res = [];
  for(let i = 1; i < data.length; i++) { 
    if(data[i][0]) {
      res.push({ date: formatDateForInput(data[i][0]), name: data[i][1] || "วันหยุด" });
    }
  }
  return res.sort((a,b) => new Date(a.date) - new Date(b.date));
}

function addHolidayAdmin(dateStr, name, actorEmail) {
  try {
    writeAuditLog(actorEmail, "Add Holiday", `เพิ่มวันหยุด: ${name} (${dateStr})`);
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Holidays");
    
    if (!sheet) { 
      sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Holidays"); 
      sheet.appendRow(["Date", "Name"]);
    }
    
    sheet.appendRow([new Date(dateStr), name]); 
    SpreadsheetApp.flush(); 
    return { success: true, message: "เพิ่มวันหยุดสำเร็จ" };
  } catch(e) { 
    return { success: false, message: e.toString() };
  }
}

function deleteHolidayAdmin(dateStr, actorEmail) {
  try {
    writeAuditLog(actorEmail, "Delete Holiday", `ลบวันหยุดวันที่: ${dateStr}`); 
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Holidays");
    if(!sheet) return { success: false, message: "ไม่พบข้อมูลแผ่นงาน" };
    
    const data = sheet.getDataRange().getValues(); 
    let target = formatDateForInput(dateStr);
    for(let i = data.length - 1; i >= 1; i--) { 
      if(formatDateForInput(data[i][0]) === target) { 
        sheet.deleteRow(i + 1);
        SpreadsheetApp.flush(); 
        return { success: true, message: "ลบสำเร็จ" }; 
      } 
    }
    return { success: false, message: "หาไม่พบวันหยุดนี้" };
  } catch(e) { 
    return { success: false, message: e.toString() };
  }
}

function getHolidays() { 
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Holidays"); 
  if(!sheet) return [];
  return sheet.getDataRange().getValues().slice(1).filter(r => r[0]).map(r => Utilities.formatDate(new Date(r[0]), Session.getScriptTimeZone(), "yyyy-MM-dd")); 
}

function calculateWorkingDays(startDateStr, endDateStr, userPosition) { 
  const holidays = getHolidays();
  const settings = getAllSettings();
  
  const weekendPosStr = settings.weekendPositions || "";
  const weekendPosArray = weekendPosStr.split(',').map(s => s.trim());
  const isWeekendWorker = userPosition && weekendPosArray.includes(userPosition);

  let workingDays = 0;
  let currentDate = new Date(startDateStr); 
  const endDate = new Date(endDateStr);
  
  while (currentDate <= endDate) { 
    let dayOfWeek = currentDate.getDay(); 
    let dateStr = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
    if (isWeekendWorker) {
      if (!holidays.includes(dateStr)) workingDays++;
    } else {
      if (dayOfWeek !== 0 && dayOfWeek !== 6 && !holidays.includes(dateStr)) {
        workingDays++;
      }
    }
    currentDate.setDate(currentDate.getDate() + 1); 
  } 
  return workingDays;
}

// ==========================================
// 📝 ระบบการลา (Leave Management)
// ==========================================
function getLeaveBalances(email, opt_nocache) {
  const settings = getAllSettings(); 
  let leaveTypes = [];
  try { 
    leaveTypes = JSON.parse(settings.leaveTypesJSON || "[]"); 
  } catch (e) { 
    leaveTypes = [ { name: 'ลาป่วย', quota: 60, icon: 'mdi-pill' } ];
  }
  
  let accumLeave = 0; 
  let empType = "อื่นๆ"; 
  const userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  if (userSheet) {
    const usersData = userSheet.getDataRange().getValues();
    for(let i = 1; i < usersData.length; i++) { 
      if(usersData[i][0] === email) { 
        accumLeave = parseFloat(usersData[i][9]) || 0; 
        empType = usersData[i][15] || "อื่นๆ"; 
        break;
      } 
    }
  }
  
  let balances = {}; 
  let currentYear = new Date().getFullYear();
  let fyStartStr = settings.fy1Start || (currentYear + "-10-01"); 
  let fyEndStr = settings.fy2End || ((currentYear + 1) + "-09-30");
  let fyStart = new Date(fyStartStr); 
  let fyEnd = new Date(fyEndStr); 
  fyEnd.setHours(23, 59, 59, 999);
  
  leaveTypes.forEach(t => { 
    let q = t.quotas ? parseInt(t.quotas[empType]) : parseInt(t.quota);
    if (isNaN(q)) q = parseInt(t.quota) || 0;
    if (t.name === 'ลาพักผ่อน') q += accumLeave; 
    balances[t.name] = { quota: q, taken: 0, remaining: q, icon: t.icon }; 
  });
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LeaveRequests"); 
  if (!sheet) return balances; 
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) { 
    let reqEmail = data[i][1];
    let type = data[i][3]; 
    let reqStart = new Date(data[i][4]); 
    let days = parseFloat(data[i][6]) || 0; 
    let status = data[i][9] || ""; 
    
    if (reqEmail === email && reqStart >= fyStart && reqStart <= fyEnd && !status.includes("ไม่อนุมัติ") && status !== "ยกเลิกการลาโดยผู้ขอ") { 
      if (balances[type]) { 
        balances[type].taken += days;
        balances[type].remaining = balances[type].quota - balances[type].taken; 
      } 
    } 
  }
  return balances;
}

function getApproversList(opt_nocache) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users"); 
  if(!sheet) return []; 
  
  const data = sheet.getDataRange().getValues(); 
  let approvers = [];
  for (let i = 1; i < data.length; i++) { 
    const roles = String(data[i][8] || 'User').split(',').map(r => r.trim());
    
    // 🚀 แก้ไข: เพิ่ม Approve2 (ปลัด) และ Approve3 (ผู้บริหาร) เข้ามาในลิสต์ด้วย
    if (roles.includes('Approve1') || roles.includes('Approve2') || roles.includes('Approve3') || roles.includes('Admin')) { 
      let aff = data[i][4] ? String(data[i][4]).trim() : "";
      let textDisplay = (data[i][1] || 'ไม่ระบุชื่อ') + " (" + (data[i][3] || "ไม่มีตำแหน่ง") + ")";
      if (aff) textDisplay += " - " + aff; 
      approvers.push({ text: textDisplay, value: data[i][0], affiliation: aff });
    } 
  } 
  return approvers;
}
function getDashboardData(userEmail, userRoles, opt_nocache) {
  let stats = { total: 0, approved: 0, pending: 0, rejected: 0 };
  let recentRequests = [];
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LeaveRequests"); 
  if (!sheet) return { stats: stats, recentRequests: recentRequests };
  const data = sheet.getDataRange().getValues(); 
  
  const canSeeAll = userRoles.includes('Admin') || userRoles.includes('Approve3') || userRoles.includes('HR');
  for (let i = data.length - 1; i >= 1; i--) { 
    const rowEmail = data[i][1];
    if (canSeeAll || rowEmail === userEmail) { 
      stats.total++; 
      const status = data[i][9] || "รอดำเนินการ";
      if (status === "อนุมัติเสร็จสิ้น") stats.approved++; 
      else if (status.includes("ไม่อนุมัติ") || status === "ยกเลิกการลาโดยผู้ขอ") stats.rejected++; 
      else stats.pending++;
      
      if (recentRequests.length < 5) { 
        recentRequests.push({ 
          date: formatDateSafe(data[i][0]), 
          fullName: data[i][2], 
          leaveType: data[i][3], 
          status: status, 
          pdfUrl: data[i][17] || "" 
        });
      } 
    } 
  } 
  return { stats: stats, recentRequests: recentRequests };
}

function getMyLeaveHistory(email, userRoles, opt_nocache) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LeaveRequests"); 
  if (!sheet) return []; 
  
  const data = sheet.getDataRange().getValues();
  let res = []; 
  
  for (let i = data.length - 1; i >= 1; i--) { 
    if (data[i][1] === email || userRoles.includes('Admin') || userRoles.includes('HR')) { 
      let formattedDateRange = formatDateSafe(data[i][4]) + " ถึง " + formatDateSafe(data[i][5]);
      let app1Info = {}, app2Info = {}, app3Info = {};
      try { app1Info = JSON.parse(data[i][19] || "{}"); } catch(e){}
      try { app2Info = JSON.parse(data[i][20] || "{}"); } catch(e){}
      try { app3Info = JSON.parse(data[i][21] || "{}"); } catch(e){}

      res.push({ 
        rowId: i + 1, 
        timestamp: Utilities.formatDate(new Date(data[i][0]), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"), 
        fullName: data[i][2], 
        leaveType: data[i][3], 
        dateRange: formattedDateRange, 
        workingDays: data[i][6], 
        status: data[i][9], 
        fileUrl: data[i][16] || "ไม่มีไฟล์แนบ", 
        pdfUrl: data[i][17] || "", 
        app1Status: data[i][10] || "รอดำเนินการ", 
        app2Status: data[i][12] || "รอดำเนินการ", 
        app3Status: data[i][14] || "รอดำเนินการ",
        app1Date: app1Info.approveDate || "", 
        app2Date: app2Info.approveDate || "", 
        app3Date: app3Info.approveDate || ""  
      });
    } 
  } 
  return res;
}

function getReportData(opt_nocache) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LeaveRequests");
  if (!sheet) return [];
  const data = sheet.getDataRange().getDisplayValues(); 
  const rawData = sheet.getDataRange().getValues();
  let results = []; 
  for (let i = 1; i < data.length; i++) { 
    results.push({ 
      timestamp: data[i][0], email: data[i][1], fullName: data[i][2], 
      leaveType: data[i][3], startDate: data[i][4], endDate: data[i][5], 
      workingDays: data[i][6], reason: data[i][7], status: data[i][9], 
      pdfUrl: rawData[i][17] || "" 
    }); 
  } 
  return results.reverse();
}

function getHREvaluationStats(startDateStr, endDateStr, opt_nocache) {
  const usersData = getAllUsersAdmin();
  let statsMap = {};
  
  usersData.forEach(u => { 
    statsMap[u.email] = { 
      fullName: u.fullName, position: u.position, 
      sickTimes: 0, sickDays: 0, personalTimes: 0, personalDays: 0, 
      vacationTimes: 0, vacationDays: 0, totalTimes: 0, totalDays: 0 
    }; 
  });
  
  let start = startDateStr ? new Date(startDateStr) : new Date(2000, 0, 1); 
  let end = endDateStr ? new Date(endDateStr) : new Date(2100, 0, 1); 
  end.setHours(23,59,59,999);
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LeaveRequests"); 
  if (!sheet) return Object.values(statsMap); 
  
  const data = sheet.getDataRange().getValues();
  for(let i = 1; i < data.length; i++) { 
    let email = data[i][1];
    let type = data[i][3]; 
    let reqStart = new Date(data[i][4]); 
    let days = parseFloat(data[i][6]) || 0; 
    let status = data[i][9] || "";
    
    if(reqStart >= start && reqStart <= end && status.includes("อนุมัติเสร็จสิ้น")) { 
      if(!statsMap[email]) continue;
      if(type.includes("ป่วย")) { statsMap[email].sickTimes++; statsMap[email].sickDays += days; } 
      else if(type.includes("กิจ")) { statsMap[email].personalTimes++; statsMap[email].personalDays += days; } 
      else if(type.includes("พักผ่อน")) { statsMap[email].vacationTimes++; statsMap[email].vacationDays += days; } 
      
      statsMap[email].totalTimes++; 
      statsMap[email].totalDays += days;
    } 
  }
  return Object.values(statsMap);
}

function isUserOnLeave(email) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LeaveRequests"); 
  if(!sheet) return false;
  const data = sheet.getDataRange().getValues(); 
  let today = new Date(); 
  today.setHours(0,0,0,0);
  for(let i=1; i<data.length; i++) { 
    if(data[i][1] === email && data[i][9].includes("อนุมัติเสร็จสิ้น")) { 
      let start = new Date(data[i][4]);
      start.setHours(0,0,0,0); 
      let end = new Date(data[i][5]); 
      end.setHours(23,59,59,999); 
      if(today >= start && today <= end) return true;
    } 
  } 
  return false;
}

function getPendingLeaves(userRoles, userEmail, opt_nocache) {
  const userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  if (!userSheet) return [];
  const usersData = userSheet.getDataRange().getValues();
  
  let myRoles = []; 
  userRoles.forEach(r => myRoles.push({ role: r, email: userEmail, delegatorName: "" }));
  let today = new Date(); 
  today.setHours(0,0,0,0);
  let userDict = {};
  
  for(let i=1; i<usersData.length; i++) { 
    let delegatorEmail = usersData[i][0];
    userDict[delegatorEmail] = { avatar: usersData[i][7] || "", position: usersData[i][3] || "" };

    let delegateEmail = usersData[i][10];
    if(delegateEmail === userEmail) { 
      let start = usersData[i][11] ? new Date(usersData[i][11]) : null;
      let end = usersData[i][12] ? new Date(usersData[i][12]) : null; 
      if (start && end) { 
        start.setHours(0,0,0,0);
        end.setHours(23,59,59,999);
        if(today >= start && today <= end) {
          let dRoles = String(usersData[i][8] || 'User').split(',').map(r => r.trim());
          dRoles.forEach(dr => { myRoles.push({ role: dr, email: delegatorEmail, delegatorName: usersData[i][1] }); });
        }
      } 
    } 
  }
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LeaveRequests");
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues(); 
  let results = [];
  
  for (let i = data.length - 1; i >= 1; i--) { 
    let status = data[i][9] || ""; 
    let targetApprove1Email = data[i][18] || ""; 
    let reqEmail = data[i][1] || ""; 
    
    if (status === "ยกเลิกการลาโดยผู้ขอ" || status.includes("ไม่อนุมัติ")) continue;
    
    let isMyTask = false;
    let actionRole = ""; 
    let delegatorNote = "";
    
    for(let r of myRoles) { 
      // 🚀 ลบเงื่อนไขการมองเห็นแบบเหมาเข่งของ Admin ออกจากคิวพิจารณา
      let targetStatus = "";
      if (r.role === "Approve1") targetStatus = "รอหัวหน้าส่วนราชการ (ผู้อนุมัติ 1)"; 
      else if (r.role === "Approve2") targetStatus = "รอปลัดพิจารณา (ผู้อนุมัติ 2)";
      else if (r.role === "Approve3") targetStatus = "รอผู้บริหารพิจารณา (ผู้อนุมัติ 3)";
      
      if (targetStatus && status.includes(targetStatus)) { 
        if (r.role === "Approve1") { 
          if (targetApprove1Email === "" || targetApprove1Email === r.email) { 
            isMyTask = true; actionRole = r.role; delegatorNote = r.delegatorName; break; 
          } 
        } else { 
          isMyTask = true; actionRole = r.role; delegatorNote = r.delegatorName; break; 
        } 
      } 
    }
    
    if (isMyTask) { 
      let formattedDateRange = formatDateSafe(data[i][4]) + " ถึง " + formatDateSafe(data[i][5]);
      let balances = getLeaveBalances(reqEmail); 
      let currentLeaveType = data[i][3];
      let remainingDays = balances[currentLeaveType] ? balances[currentLeaveType].remaining : 0;
      
      let app1Info = {}, app2Info = {}, app3Info = {};
      try { app1Info = JSON.parse(data[i][19] || "{}"); } catch(e){}
      try { app2Info = JSON.parse(data[i][20] || "{}"); } catch(e){}
      try { app3Info = JSON.parse(data[i][21] || "{}"); } catch(e){}

      results.push({ 
        rowId: i + 1, timestamp: Utilities.formatDate(new Date(data[i][0]), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"), 
        fullName: data[i][2], leaveType: data[i][3], dateRange: formattedDateRange, 
        workingDays: data[i][6], reason: data[i][7], status: data[i][9],
        remainingDays: remainingDays, fileUrl: data[i][16] || "ไม่มีไฟล์แนบ", 
        actionRole: actionRole, delegatorNote: delegatorNote,
        avatar: userDict[reqEmail] ? userDict[reqEmail].avatar : "",
        position: userDict[reqEmail] ? userDict[reqEmail].position : "",
        app1Status: data[i][10] || "รอดำเนินการ", app2Status: data[i][12] || "รอดำเนินการ", app3Status: data[i][14] || "รอดำเนินการ",
        app1Date: app1Info.approveDate || "", app2Date: app2Info.approveDate || "", app3Date: app3Info.approveDate || ""
      });
    } 
  }
  return results;
}

function submitLeaveRequest(data) {
  try {
    writeAuditLog(data.email, "Submit Leave", `ยื่น${data.leaveType} (${data.workingDays} วัน)`);
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("LeaveRequests");
    
    if(!sheet) { 
      sheet = ss.insertSheet("LeaveRequests");
      sheet.appendRow(["Timestamp", "Email", "FullName", "LeaveType", "StartDate", "EndDate", "WorkingDays", "Reason", "RequesterSign", "Status", "Approve1", "App1Sign", "Approve2", "App2Sign", "Approve3", "App3Sign", "AttachmentURL", "PDF_URL", "TargetApprove1_Email", "App1Info", "App2Info", "App3Info"]);
    }
    
    let fileUrl = "ไม่มีไฟล์แนบ";
    if (data.fileBase64 && data.fileName) { 
      const settings = getAllSettings();
      if(settings.folderId) { 
        const folderId = cleanFolderId(settings.folderId); 
        try {
          const folder = DriveApp.getFolderById(folderId);
          const contentType = data.fileBase64.substring(5, data.fileBase64.indexOf(';')); 
          const bytes = Utilities.base64Decode(data.fileBase64.split(',')[1]); 
          fileUrl = folder.createFile(Utilities.newBlob(bytes, contentType, "แนบ_" + data.fullName + "_" + data.fileName)).getUrl();
        } catch (e) { console.error("File Upload Error:", e); }
      } 
    }

    // 🚀 1. ค้นหาว่าคนที่ถูกเลือกใน Dropdown มีสิทธิ์ระดับไหน (Approve 1, 2 หรือ 3)
    let targetRoles = ['Approve1'];
    const userSheet = ss.getSheetByName("Users");
    if (userSheet) {
      const uData = userSheet.getDataRange().getValues();
      for (let i = 1; i < uData.length; i++) {
        if (uData[i][0] === data.approver1Email) { // ดึงอีเมลคนที่ถูกเลือก
          targetRoles = String(uData[i][8] || 'User').split(',').map(r => r.trim());
          break;
        }
      }
    }

    // 🚀 2. กำหนดสถานะการวิ่งของเอกสาร ตามตำแหน่งของคนที่ถูกเลือก
    let initialStatus = "รอหัวหน้าส่วนราชการ (ผู้อนุมัติ 1)";
    let app1Status = "รอดำเนินการ";
    let app2Status = "รอดำเนินการ";

    if (targetRoles.includes("Approve3") || targetRoles.includes("Admin")) {
      // ถ้าเลือก นายก (Approve 3) -> วิ่งตรงไปหานายก ข้ามขั้น 1 และ 2
      initialStatus = "รอผู้บริหารพิจารณา (ผู้อนุมัติ 3)";
      app1Status = "ข้าม (ส่งเรื่องตรงถึงผู้บริหาร)";
      app2Status = "ข้าม (ส่งเรื่องตรงถึงผู้บริหาร)";
    } else if (targetRoles.includes("Approve2")) {
      // ถ้าเลือก ปลัด (Approve 2) -> วิ่งไปหาปลัดทันที ข้ามขั้น 1
      initialStatus = "รอปลัดพิจารณา (ผู้อนุมัติ 2)";
      app1Status = "ข้าม (ส่งเรื่องตรงถึงปลัด)";
    }
    
    sheet.appendRow([ new Date(), data.email, data.fullName, data.leaveType, data.startDate, data.endDate, data.workingDays, data.reason, data.signature, initialStatus, app1Status, "", app2Status, "", "รอดำเนินการ", "", fileUrl, "", data.approver1Email, "", "", "" ]);
    SpreadsheetApp.flush(); 
    
    // 🚀 3. แจ้งเตือนเข้า Telegram ไปหาคนที่ถูกเลือก
    try {
      const sender = getUserInfoByEmail(data.email);
      const targetApprover = getUserInfoByEmail(data.approver1Email);
      const attachUrl = fileUrl === "ไม่มีไฟล์แนบ" ? "" : fileUrl;
      
      if (targetApprover.tgChatId) {
        let msgToApprover = `🔔 <b>มีคำขออนุมัติใบลาใหม่</b>\n\n👤 <b>ผู้ขอลา:</b> ${sender.name}\n📅 <b>ประเภทการลา:</b> ${data.leaveType} (${data.workingDays} วัน)\n📆 <b>วันที่:</b> ${data.startDate} ถึง ${data.endDate}\n📝 <b>เหตุผล:</b> ${data.reason}\n-------------------------\n⚠️ <b>เรียน คุณ ${targetApprover.name}</b>\nกรุณาเข้าสู่ระบบเพื่อพิจารณาอนุมัติใบลาครับ`;
        sendTelegram(msgToApprover, attachUrl, "", targetApprover.tgChatId);
      } else {
        let msgFallback = `🔔 <b>มีคำขออนุมัติใบลาใหม่</b>\n👤 <b>ผู้ขอ:</b> ${data.fullName}\n📌 <b>ประเภท:</b> ${data.leaveType}\n📅 <b>วันที่:</b> ${data.startDate} ถึง ${data.endDate}\n⏳ <b>จำนวน:</b> ${data.workingDays} วัน`;
        sendTelegram(msgFallback, attachUrl, "");
      }

      if (sender.tgChatId) {
        let msgToSender = `📄 <b>ระบบได้รับใบลาของคุณแล้ว</b>\n\n📌 <b>ประเภทการลา:</b> ${data.leaveType} (${data.workingDays} วัน)\n📍 <b>สถานะปัจจุบัน:</b> ส่งเรื่องถึง <b>คุณ ${targetApprover.name}</b> เพื่อพิจารณา\n<i>(ระบบจะแจ้งความคืบหน้าให้ทราบในข้อความถัดไป)</i>`;
        sendTelegram(msgToSender, "", "", sender.tgChatId);
      }
    } catch (tgError) { console.error("Telegram Error:", tgError); }
    
    return { success: true, message: "ยื่นใบลาเรียบร้อยแล้ว ระบบกำลังส่งเรื่องให้ผู้พิจารณา" };
  } catch (error) { 
    return { success: false, message: "Error: " + error.toString() };
  }
}

function processLeaveApproval(rowId, actionRole, decision, approverSign, approverName, approverPosition, delegatorNote, actorEmail, comment = "") {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LeaveRequests");
    const rowData = sheet.getRange(rowId, 1, 1, sheet.getLastColumn()).getValues()[0];
    let currentStatus = rowData[9];
    let nextStatus = "";
    
    writeAuditLog(actorEmail, "Approval", `พิจารณาใบลาแถวที่ ${rowId} (${decision}) ความเห็น: ${comment}`);
    
    let isCancelFlow = currentStatus.includes("ขอยกเลิก");
    if (isCancelFlow) {
      if (actionRole === "Approve1") nextStatus = (decision === "อนุมัติ") ? "ขอยกเลิก - รอปลัดพิจารณา (ผู้อนุมัติ 2)" : "อนุมัติเสร็จสิ้น"; 
      else if (actionRole === "Approve2") nextStatus = (decision === "อนุมัติ") ? "ขอยกเลิก - รอผู้บริหารพิจารณา (ผู้อนุมัติ 3)" : "อนุมัติเสร็จสิ้น"; 
      else if (actionRole === "Approve3") nextStatus = (decision === "อนุมัติ") ? "ยกเลิกการลาโดยผู้ขอ" : "อนุมัติเสร็จสิ้น"; 
      
      sheet.getRange(rowId, 10).setValue(nextStatus); 
      SpreadsheetApp.flush();
      
      const sender = getUserInfoByEmail(rowData[1]);
      let cancelMsg = `🔔 <b>แจ้งสถานะการขอยกเลิกใบลา</b>\n👤 <b>ผู้ขอ:</b> ${rowData[2]}\n✅ <b>ผลการพิจารณา:</b> ${decision}\n📍 <b>สถานะล่าสุด:</b> ${nextStatus}\n💬 <b>ความเห็น:</b> ${comment || "-"}`;
      sendTelegram(cancelMsg, "", "", sender.tgChatId);
      return {success: true, message: "บันทึกผลการขอยกเลิกเรียบร้อยแล้ว"};
    }
    
    let leaveType = rowData[3];
    let workingDays = parseFloat(rowData[6]) || 0; 
    let isBypassToFinishApprove1 = false, isBypassToFinishApprove2 = false; 

    try {
      let appConfig = getAppConfig();
      let flowRules = appConfig.flowRules || [];
      for (let rule of flowRules) {
        let ruleMaxDays = parseFloat(rule.maxDays) || 999; 
        if (leaveType === rule.leaveType && workingDays <= ruleMaxDays) {
          if (String(rule.finishStep) === "1") isBypassToFinishApprove1 = true;
          if (String(rule.finishStep) === "2") isBypassToFinishApprove2 = true;
          break; 
        }
      }
    } catch (e) { console.error("Flow rules err: " + e); }

    let approveDateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
    let signerInfo = JSON.stringify({ 
      name: approverName || "ไม่ระบุ", position: approverPosition || "", delegator: delegatorNote || "", approveDate: approveDateStr, comment: comment 
    });
    
    if (actionRole === "Approve1") { 
      sheet.getRange(rowId, 11).setValue(decision);
      sheet.getRange(rowId, 12).setValue(approverSign); 
      sheet.getRange(rowId, 20).setValue(signerInfo);
      nextStatus = (decision === "อนุมัติ") ? (isBypassToFinishApprove1 ? "อนุมัติเสร็จสิ้น" : "รอปลัดพิจารณา (ผู้อนุมัติ 2)") : "ไม่อนุมัติโดยหัวหน้าส่วนราชการ";
    } else if (actionRole === "Approve2") { 
      sheet.getRange(rowId, 13).setValue(decision);
      sheet.getRange(rowId, 14).setValue(approverSign); 
      sheet.getRange(rowId, 21).setValue(signerInfo);
      nextStatus = (decision === "อนุมัติ") ? (isBypassToFinishApprove2 ? "อนุมัติเสร็จสิ้น" : "รอผู้บริหารพิจารณา (ผู้อนุมัติ 3)") : "ไม่อนุมัติโดยปลัด";
    } else if (actionRole === "Approve3") { 
      sheet.getRange(rowId, 15).setValue(decision);
      sheet.getRange(rowId, 16).setValue(approverSign); 
      sheet.getRange(rowId, 22).setValue(signerInfo);
      nextStatus = (decision === "อนุมัติ") ? "อนุมัติเสร็จสิ้น" : "ไม่อนุมัติโดยผู้บริหาร";
    }
    
    sheet.getRange(rowId, 10).setValue(nextStatus); 
    SpreadsheetApp.flush();
    
    const requester = getUserInfoByEmail(rowData[1]);
    const currentActor = getUserInfoByEmail(actorEmail); 
    const fileUrl = rowData[16];
    
    if (decision === "อนุมัติ") {
      if (nextStatus === "อนุมัติเสร็จสิ้น") {
        generateOfficialPDF(rowId);
        try { if(typeof addLeaveToCalendar === "function") addLeaveToCalendar(rowId); } catch(e){} 
        const pdfUrl = sheet.getRange(rowId, 18).getValue(); 
        
        let msgFinal = `🎉 <b>ใบลาของคุณได้รับการอนุมัติสมบูรณ์แล้ว!</b>\n\n📅 <b>ประเภท:</b> ${rowData[3]}\n✅ อนุมัติโดย: <b>คุณ ${currentActor.name}</b>\n💬 <b>ความเห็น:</b> ${comment || "-"}\n-------------------------\nพนักงานสามารถดาวน์โหลดใบลา PDF ได้จากปุ่มด้านล่างครับ`;
        sendTelegram(msgFinal, "", pdfUrl, requester.tgChatId); 
      } else {
        let nextApproverEmail = (actionRole === "Approve1") ? rowData[18] : ((actionRole === "Approve2") ? rowData[19] : ""); 
        const nextApprover = getUserInfoByEmail(nextApproverEmail);
        
        if (nextApprover && nextApprover.tgChatId) {
          let msgToNext = `✅ <b>ใบลาผ่านการพิจารณาขั้นก่อนหน้าแล้ว</b>\n\n👤 <b>ผู้ขอลา:</b> ${requester.name}\n🟢 <b>ผ่านการอนุมัติโดย:</b> คุณ ${currentActor.name}\n💬 <b>ความเห็น/บันทึก:</b> ${comment || "-"}\n-------------------------\n⚠️ <b>เรียน คุณ ${nextApprover.name}</b>\nกรุณาพิจารณาอนุมัติใบลาในลำดับถัดไปครับ`;
          sendTelegram(msgToNext, fileUrl, "", nextApprover.tgChatId);
        }
        let msgToReq = `⏳ <b>อัปเดตสถานะใบลาของคุณ</b>\n\n✅ <b>คุณ ${currentActor.name}</b> พิจารณาอนุมัติแล้ว\n💬 <b>ความเห็น:</b> ${comment || "-"}\n📍 <b>สถานะปัจจุบัน:</b> ${nextStatus}`;
        sendTelegram(msgToReq, "", "", requester.tgChatId);
      }
    } else {
      let msgReject = `❌ <b>แจ้งผลการพิจารณาใบลา (ไม่อนุมัติ)</b>\n\nเรียน คุณ ${requester.name}\nใบลาประเภท ${rowData[3]} ของคุณไม่ได้รับการอนุมัติ\n\n🚫 <b>พิจารณาโดย:</b> คุณ ${currentActor.name}\n💬 <b>เหตุผล:</b> ${comment || "ไม่ได้ระบุเหตุผล"}`;
      sendTelegram(msgReject, "", "", requester.tgChatId);
    }
    return {success: true, message: "บันทึกผลเรียบร้อยแล้ว"};
  } catch(e) { 
    return {success: false, message: e.toString()};
  }
}

function cancelLeaveRequest(rowId, userEmail, comment = "") {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LeaveRequests");
    const rowData = sheet.getRange(rowId, 1, 1, sheet.getLastColumn()).getValues()[0];
    const emailInSheet = rowData[1], currentStatus = rowData[9], leaveType = rowData[3], approver1Email = rowData[18];
    
    if (emailInSheet !== userEmail) return { success: false, message: "คุณไม่มีสิทธิ์ยกเลิกใบลาของผู้อื่น" };
    if (currentStatus.includes("ยกเลิก")) return { success: false, message: "ใบนี้ถูกยกเลิกแล้ว" };
    
    let nextStatus = "ยกเลิกการลาโดยผู้ขอ", msg = "ยกเลิกใบลาและคืนโควตาเรียบร้อยแล้ว", isCancelAfterApprove = false; 
    
    if (currentStatus === "อนุมัติเสร็จสิ้น") { 
      nextStatus = "ขอยกเลิก - รอหัวหน้าส่วนราชการ (ผู้อนุมัติ 1)";
      msg = "ส่งคำขอยกเลิกให้พิจารณาตามลำดับแล้ว"; 
      isCancelAfterApprove = true;
    }
    
    sheet.getRange(rowId, 10).setValue(nextStatus); 
    SpreadsheetApp.flush();
    writeAuditLog(userEmail, "Cancel Leave", `ขอยกเลิกใบลาแถวที่ ${rowId} (${nextStatus}) เหตุผล: ${comment}`);
    
    const requester = getUserInfoByEmail(emailInSheet);
    if (isCancelAfterApprove) {
      const approver1 = getUserInfoByEmail(approver1Email);
      if (approver1.tgChatId) {
        let msgToApp1 = `⚠️ <b>มีคำขอยกเลิกใบลา (รออนุมัติขั้นที่ 1)</b>\n\n👤 <b>ผู้ขอ:</b> ${requester.name}\n📌 <b>ใบลาที่ขอยกเลิก:</b> ${leaveType}\n💬 <b>เหตุผล:</b> ${comment || "-"}\n-------------------------\nกรุณาพิจารณาด้วยครับ`;
        sendTelegram(msgToApp1, "", "", approver1.tgChatId);
      }
    } else {
      if (requester.tgChatId) {
        let msgToReq = `🗑️ <b>ยกเลิกคำขอลาสำเร็จ</b>\n\nใบลาประเภท <b>${leaveType}</b> ถูกยกเลิก\n💬 <b>เหตุผล:</b> ${comment || "-"}\n✅ ระบบคืนโควตาวันลาเรียบร้อยครับ`;
        sendTelegram(msgToReq, "", "", requester.tgChatId);
      }
    }
    return { success: true, message: msg };
  } catch (e) { 
    return { success: false, message: e.toString() };
  }
}

function generateOfficialPDF(rowId) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LeaveRequests"); 
    const data = sheet.getRange(rowId, 1, 1, 22).getValues()[0];
    const settings = getAllSettings(); 
    const folderId = cleanFolderId(settings.folderId); 
    if (!folderId) return; 
    
    let position = "", affiliation = "";
    const userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
    if (userSheet) {
      const users = userSheet.getDataRange().getValues();
      for(let i=1; i<users.length; i++) { 
        if(users[i][0] === data[1]) { 
          position = users[i][3] || ""; affiliation = users[i][4] || ""; break; 
        } 
      }
    }
    
    let logoData = settings.logo || ""; 
    if(!logoData || logoData.length < 100) { 
      try { 
        let res = UrlFetchApp.fetch("https://raw.githubusercontent.com/tientn/Garuda/master/garuda.png");
        logoData = "data:image/png;base64," + Utilities.base64Encode(res.getBlob().getBytes()); 
      } catch(e) { logoData = ""; } 
    }
    
    let app1Info = {name: "ไม่ระบุ", position: "หัวหน้าส่วนราชการ", delegator: ""};
    let app2Info = {name: "ไม่ระบุ", position: "ปลัด อบต.", delegator: ""}; 
    let app3Info = {name: "ไม่ระบุ", position: "นายกองค์การบริหารส่วนตำบล", delegator: ""};
    try { if(data[19]) app1Info = JSON.parse(data[19]); } catch(e){} 
    try { if(data[20]) app2Info = JSON.parse(data[20]); } catch(e){} 
    try { if(data[21]) app3Info = JSON.parse(data[21]); } catch(e){}
    
    const htmlTemplate = HtmlService.createTemplateFromFile('PdfTemplate');
    htmlTemplate.data = { 
      schoolName: settings.schoolName || "องค์การบริหารส่วนตำบล", 
      logoData: logoData, 
      timestamp: Utilities.formatDate(new Date(data[0]), Session.getScriptTimeZone(), "dd/MM/yyyy"), 
      fullName: data[2], leaveType: data[3], startDate: formatDateSafe(data[4]), 
      endDate: formatDateSafe(data[5]), workingDays: data[6], reason: data[7], 
      reqSign: data[8] ? String(data[8]).trim() : "", 
      app1Sign: data[11] ? String(data[11]).trim() : "", 
      app2Sign: data[13] ? String(data[13]).trim() : "", 
      app3Sign: data[15] ? String(data[15]).trim() : "", 
      app1Info: app1Info, app2Info: app2Info, app3Info: app3Info, 
      position: position, affiliation: affiliation 
    };
    
    const htmlOutput = htmlTemplate.evaluate(); 
    const blob = htmlOutput.getAs('application/pdf').setName("ใบลา_" + data[2] + ".pdf");
    const folder = DriveApp.getFolderById(folderId); 
    const pdfFile = folder.createFile(blob);
    // เปิดสิทธิ์ไฟล์ให้คนในลิงก์อ่านได้
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    sheet.getRange(rowId, 18).setValue(pdfFile.getUrl()); 
    SpreadsheetApp.flush();
  } catch (e) { 
    console.error("PDF Error: " + e);
  }
}
// ==========================================
// 🏢 ระบบจองห้องประชุม 
// ==========================================
function getRoomsAdmin(opt_nocache) {
  try {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Rooms");
    if (!sheet) { 
      sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Rooms"); 
      sheet.appendRow(["RoomName", "Capacity", "Facilities", "Color", "Status", "ImageUrl"]);
      sheet.appendRow(["ห้องประชุมสภา", "50", "โปรเจคเตอร์, ไมค์, แอร์", "#3b82f6", "เปิดใช้งาน", ""]); 
    }
    
    if (sheet.getLastColumn() < 6) {
      sheet.getRange(1, 6).setValue("ImageUrl");
    }
    
    const data = sheet.getDataRange().getValues(); 
    let res = [];
    for(let i=1; i<data.length; i++) { 
      if(data[i][0]) {
        res.push({ 
          name: data[i][0], 
          capacity: data[i][1] || "0", 
          facilities: data[i][2] || "-", 
          color: data[i][3] || "#cccccc", 
          status: data[i][4] || "เปิดใช้งาน",
          imageUrl: data[i][5] || "" 
        });
      }
    } 
    return res;
  } catch (e) {
    console.error("getRoomsAdmin Error: " + e);
    return [];
  }
}

function saveRoomAdmin(roomObj, actorEmail) {
  try { 
    writeAuditLog(actorEmail, "Room Setup", `อัปเดตห้องประชุม: ${roomObj.name}`);
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Rooms"); 
    
    if (!sheet) { 
      sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Rooms");
      sheet.appendRow(["RoomName", "Capacity", "Facilities", "Color", "Status", "ImageUrl"]); 
    } 
    
    const data = sheet.getDataRange().getValues();
    for(let i=1; i<data.length; i++) { 
      if(data[i][0] === roomObj.name) { 
        sheet.getRange(i+1, 2, 1, 5).setValues([[roomObj.capacity, roomObj.facilities, roomObj.color, roomObj.status, roomObj.imageUrl || ""]]);
        SpreadsheetApp.flush(); 
        return {success:true, message:"บันทึกการแก้ไขห้องสำเร็จ"}; 
      } 
    } 
    
    sheet.appendRow([roomObj.name, roomObj.capacity, roomObj.facilities, roomObj.color, roomObj.status || "เปิดใช้งาน", roomObj.imageUrl || ""]);
    SpreadsheetApp.flush(); 
    return {success:true, message:"เพิ่มห้องประชุมใหม่สำเร็จ"}; 
  } catch(e) { 
    return {success:false, message:e.toString()};
  }
}

function deleteRoomAdmin(roomName, actorEmail) {
  try { 
    writeAuditLog(actorEmail, "Room Delete", `ลบห้องประชุม: ${roomName}`); 
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Rooms");
    if (!sheet) return { success: false, message: "ไม่พบฐานข้อมูล" };
    
    const data = sheet.getDataRange().getValues(); 
    for(let i = data.length-1; i >= 1; i--) { 
      if(data[i][0] === roomName) { 
        sheet.deleteRow(i+1);
        SpreadsheetApp.flush(); 
        return {success:true, message:"ลบห้องประชุมสำเร็จ"}; 
      } 
    } 
    return {success:false, message:"ไม่พบห้องนี้"};
  } catch(e) { 
    return {success:false, message:e.toString()}; 
  }
}

function checkRoomAvailability(roomName, startDtStr, endDtStr) {
  try {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("RoomBookings");
    if(!sheet) return true; 
    
    const reqStart = new Date(startDtStr).getTime(); 
    const reqEnd = new Date(endDtStr).getTime(); 
    const data = sheet.getDataRange().getValues();
    
    for(let i=1; i<data.length; i++) { 
      if(data[i][4] === roomName && data[i][10] !== "ไม่อนุมัติ" && data[i][10] !== "ยกเลิก") { 
        let bookStart = new Date(data[i][5]).getTime();
        let bookEnd = new Date(data[i][6]).getTime(); 
        if (reqStart < bookEnd && reqEnd > bookStart) { 
          return false; // เวลาชนกัน
        } 
      } 
    } 
    return true;
  } catch(e) {
    return true;
  }
}

function submitRoomBooking(data) {
  try {
    writeAuditLog(data.email, "Room Booking", `จองห้อง: ${data.roomName}`); 
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("RoomBookings");
    if(!sheet) { 
      sheet = ss.insertSheet("RoomBookings");
      sheet.appendRow(["BookingID", "Timestamp", "Email", "FullName", "RoomName", "StartDateTime", "EndDateTime", "Title", "Reason", "Attendees", "Status", "ApproverSign", "CalendarEventId", "Approve1Name", "Approve2Name"]);
    }
    if(sheet.getLastColumn() >= 14 && sheet.getRange(1, 14).getValue() !== "Approve1Name") { 
      sheet.getRange(1, 14).setValue("Approve1Name");
      sheet.getRange(1, 15).setValue("Approve2Name"); 
    }
    if(!checkRoomAvailability(data.roomName, data.startDateTime, data.endDateTime)) { 
      return { success: false, message: "ขออภัย! มีผู้อื่นทำรายการจองห้องนี้ในช่วงเวลาดังกล่าวตัดหน้าไปแล้ว" };
    }
    const bookingId = "RB" + new Date().getTime();
    const initialStatus = "รออนุมัติ 1 (Pending)";
    sheet.appendRow([bookingId, new Date(), data.email, data.fullName, data.roomName, data.startDateTime, data.endDateTime, data.title, data.reason, data.attendees, initialStatus, "", "", "", ""]);
    SpreadsheetApp.flush();
    
    let tgMsg = `🏢 <b>มีการขอจองห้องประชุมใหม่</b>\n👤 <b>ผู้จอง:</b> ${cleanTgText(data.fullName)}\n🚪 <b>ห้อง:</b> ${cleanTgText(data.roomName)}\n📌 <b>หัวข้อ:</b> ${cleanTgText(data.title)}\n📅 <b>เวลา:</b> ${formatDateTimeSafe(data.startDateTime)} ถึง ${formatDateTimeSafe(data.endDateTime)}\n⏳ <b>สถานะ:</b> รอผู้อนุมัติห้องประชุม (ขั้นที่ 1)`;
    sendTelegram(tgMsg, "", "", "", "room"); // 🚀 บังคับให้ใช้บอทของห้องประชุม
    
    return { success: true, message: "ส่งคำขอจองห้องประชุมเรียบร้อยแล้ว" };
  } catch (e) { 
    return { success: false, message: "Error: " + e.toString() };
  }
}

function getMyRoomBookings(email, userRoles, opt_nocache) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("RoomBookings"); 
  if (!sheet) return []; 
  
  const data = sheet.getDataRange().getValues();
  let results = [];
  
  for (let i = data.length - 1; i >= 1; i--) { 
    if (data[i][2] === email || userRoles.includes('Admin') || userRoles.includes('RoomApprove1') || userRoles.includes('RoomApprove2')) { 
      results.push({ 
        rowId: i + 1, 
        bookingId: data[i][0], 
        timestamp: formatDateTimeSafe(data[i][1]), 
        email: data[i][2], 
        fullName: data[i][3], 
        roomName: data[i][4], 
        startDateTime: formatDateTimeSafe(data[i][5]), 
        endDateTime: formatDateTimeSafe(data[i][6]), 
        title: data[i][7], 
        reason: data[i][8], 
        attendees: data[i][9], 
        status: data[i][10] 
      });
    } 
  } 
  return results;
}

function getPendingRoomBookings(userRoles, opt_nocache) {
  if (!userRoles.some(r => ['Admin', 'RoomApprove1', 'RoomApprove2'].includes(r))) return [];
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("RoomBookings"); 
  if (!sheet) return []; 
  
  const data = sheet.getDataRange().getValues(); 
  let results = [];
  for (let i = data.length - 1; i >= 1; i--) {
    let status = data[i][10];
    let isMyTask = false;
    
    // 🚀 ป้องกันแอดมินที่ไม่มี Role เฉพาะ มองเห็นคิวงานห้องประชุม
    if (userRoles.includes('RoomApprove1') && status === "รออนุมัติ 1 (Pending)") isMyTask = true;
    if (userRoles.includes('RoomApprove2') && status === "รออนุมัติ 2 (Pending)") isMyTask = true;
    
    if (isMyTask) { 
      results.push({ 
        rowId: i + 1, 
        bookingId: data[i][0], 
        timestamp: formatDateTimeSafe(data[i][1]), 
        fullName: data[i][3], 
        roomName: data[i][4], 
        startDateTime: formatDateTimeSafe(data[i][5]), 
        endDateTime: formatDateTimeSafe(data[i][6]), 
        title: data[i][7], 
        reason: data[i][8], 
        attendees: data[i][9], 
        status: data[i][10] 
      });
    }
  }
  return results;
}

function approveRoomBooking(rowId, decision, approverName, email, comment = "") {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("RoomBookings");
    writeAuditLog(email, "Room Approval", `พิจารณาห้องประชุมแถว ${rowId} -> ${decision} (${comment})`);
    let currentStatus = sheet.getRange(rowId, 11).getValue(); 
    let nextStatus = "";
    if (decision === "ไม่อนุมัติ") { nextStatus = "ไม่อนุมัติ"; } 
    else {
      if (currentStatus === "รออนุมัติ 1 (Pending)") { nextStatus = "รออนุมัติ 2 (Pending)"; sheet.getRange(rowId, 14).setValue(approverName); } 
      else if (currentStatus === "รออนุมัติ 2 (Pending)") { nextStatus = "อนุมัติแล้ว (Approved)"; sheet.getRange(rowId, 15).setValue(approverName); } 
      else { nextStatus = "อนุมัติแล้ว (Approved)"; }
    }
    sheet.getRange(rowId, 11).setValue(nextStatus);
    
    if(nextStatus === "อนุมัติแล้ว (Approved)") {
      const data = sheet.getRange(rowId, 1, 1, 10).getValues()[0];
      const settings = getAllSettings(); 
      const calId = settings.roomCalendarId ? String(settings.roomCalendarId).trim() : String(settings.calendarId || "").trim();
      if(calId) { try { const cal = CalendarApp.getCalendarById(calId); if(cal) { const eventTitle = `[ห้องประชุม] ${data[4]} - ${data[7]} (${data[3]})`; const ev = cal.createEvent(eventTitle, new Date(data[5]), new Date(data[6]), {description: data[8]}); sheet.getRange(rowId, 13).setValue(ev.getId()); } } catch(e) {} }
    }
    SpreadsheetApp.flush();
    
    const roomName = sheet.getRange(rowId, 5).getValue(); 
    const reqName = sheet.getRange(rowId, 4).getValue(); 
    const reqEmail = sheet.getRange(rowId, 3).getValue();
    
    let tgMsg = `🔔 <b>อัปเดตสถานะห้องประชุม</b>\n👤 ผู้จอง: ${cleanTgText(reqName)}\n🚪 ห้อง: ${cleanTgText(roomName)}\n✅ <b>สถานะล่าสุด: ${cleanTgText(nextStatus)}</b>\n💬 <b>ความเห็น:</b> ${cleanTgText(comment || "-")}`;
    let tgId = getUserPersonalTgChatId(reqEmail); 
    if(tgId) sendTelegram(tgMsg, "", "", tgId, "room"); // 🚀 บังคับให้ใช้บอทของห้องประชุม
    sendTelegram(tgMsg, "", "", "", "room"); // 🚀 บังคับให้ใช้บอทของห้องประชุม
    
    return { success: true, message: "บันทึกผลสำเร็จ" };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function cancelRoomBooking(rowId, email, comment = "") {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("RoomBookings");
    if(sheet.getRange(rowId, 3).getValue() !== email) return {success:false, message:"ไม่มีสิทธิ์ยกเลิกรายการนี้"};
    sheet.getRange(rowId, 11).setValue("ยกเลิก");
    const eventId = sheet.getRange(rowId, 13).getValue();
    if(eventId) { 
      const settings = getAllSettings(); 
      const calId = settings.roomCalendarId || settings.calendarId;
      try { let cal = CalendarApp.getCalendarById(calId); if(cal) cal.getEventById(eventId).deleteEvent(); } catch(e){} 
    }
    writeAuditLog(email, "Cancel Room", `ยกเลิกจองห้องแถว ${rowId} เหตุผล: ${comment}`);
    SpreadsheetApp.flush();
    return {success:true, message:"ยกเลิกการจองสำเร็จ"};
  } catch(e) { return {success:false, message:e.toString()}; }
}

function approveVehicleBooking(rowId, decision, approverName, email, comment = "") {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VehicleBookings");
    writeAuditLog(email, "Vehicle Approval", `พิจารณาจองรถแถว ${rowId} -> ${decision} (${comment})`);
    let nextStatus = decision === "ไม่อนุมัติ" ? "ไม่อนุมัติ" : "อนุมัติแล้ว (Approved)";
    sheet.getRange(rowId, 11).setValue(nextStatus);
    if(nextStatus === "อนุมัติแล้ว (Approved)") sheet.getRange(rowId, 12).setValue(approverName);
    
    if(nextStatus === "อนุมัติแล้ว (Approved)") {
      const data = sheet.getRange(rowId, 1, 1, 10).getValues()[0];
      const settings = getAllSettings(); 
      const calId = settings.vehicleCalendarId ? String(settings.vehicleCalendarId).trim() : String(settings.calendarId || "").trim();
      if(calId) { 
        try { 
          const cal = CalendarApp.getCalendarById(calId);
          if(cal) { 
            const eventTitle = `[จองรถ] ${data[4]} - ${data[7]} (${data[3]})`;
            const ev = cal.createEvent(eventTitle, new Date(data[5]), new Date(data[6]), {description: data[8]});
            sheet.getRange(rowId, 13).setValue(ev.getId());
          } 
        } catch(e) {} 
      }
    }
    SpreadsheetApp.flush();
    const licensePlate = sheet.getRange(rowId, 5).getValue(); 
    const reqName = sheet.getRange(rowId, 4).getValue(); 
    const reqEmail = sheet.getRange(rowId, 3).getValue();
    
    let tgMsg = `🔔 <b>อัปเดตสถานะจองรถ</b>\n👤 ผู้จอง: ${cleanTgText(reqName)}\n🚘 รถ: ${cleanTgText(licensePlate)}\n✅ <b>สถานะล่าสุด: ${cleanTgText(nextStatus)}</b>\n💬 <b>ความเห็น:</b> ${cleanTgText(comment || "-")}`;
    let tgId = getUserPersonalTgChatId(reqEmail); 
    if(tgId) sendTelegram(tgMsg, "", "", tgId); 
    sendTelegram(tgMsg);
    
    return { success: true, message: "บันทึกผลสำเร็จ" };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function cancelVehicleBooking(rowId, email, comment = "") {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VehicleBookings");
    if(sheet.getRange(rowId, 3).getValue() !== email) return {success:false, message:"ไม่มีสิทธิ์ยกเลิกรายการนี้"};
    sheet.getRange(rowId, 11).setValue("ยกเลิก");
    const eventId = sheet.getRange(rowId, 13).getValue();
    if(eventId) { 
      const settings = getAllSettings(); 
      const calId = settings.vehicleCalendarId || settings.calendarId;
      try { let cal = CalendarApp.getCalendarById(calId); if(cal) cal.getEventById(eventId).deleteEvent(); } catch(e){} 
    }
    writeAuditLog(email, "Cancel Vehicle", `ยกเลิกจองรถแถว ${rowId} เหตุผล: ${comment}`);
    SpreadsheetApp.flush();
    return {success:true, message:"ยกเลิกการจองสำเร็จ"};
  } catch(e) { return {success:false, message:e.toString()}; }
}

// ==========================================
// 🚗 ระบบจัดการรถส่วนกลาง (Vehicles & Maintenance)
// ==========================================

function getVehiclesAdmin(opt_nocache) {
  try {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Vehicles");
    if (!sheet) {
      sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Vehicles");
      sheet.appendRow(["LicensePlate", "BrandModel", "Type", "TaxDueDate", "SmokeDueDate", "ImageUrl", "ResponsibleEmail"]);
      return [];
    }
    
    if (sheet.getLastColumn() < 7) {
      sheet.getRange(1, 7).setValue("ResponsibleEmail");
    }
    
    const data = sheet.getDataRange().getValues();
    let res = [];
    for(let i=1; i<data.length; i++) {
      if(data[i][0]) {
        res.push({
          rowId: i + 1,
          licensePlate: data[i][0],
          brandModel: data[i][1],
          type: data[i][2],
          taxDueDate: formatDateForInput(data[i][3]),
          smokeDueDate: data[i][4] ? formatDateForInput(data[i][4]) : "",
          imageUrl: data[i][5] || "",
          responsibleEmail: data[i][6] || ""
        });
      }
    }
    return res;
  } catch (e) {
    return [];
  }
}

function saveVehicleAdmin(vehicleObj, actorEmail) {
  try {
    writeAuditLog(actorEmail, "Vehicle Setup", `อัปเดตข้อมูลรถ: ${vehicleObj.licensePlate}`);
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Vehicles");
    if (!sheet) {
      sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Vehicles");
      sheet.appendRow(["LicensePlate", "BrandModel", "Type", "TaxDueDate", "SmokeDueDate", "ImageUrl", "ResponsibleEmail"]);
    }
    
    let smokeDateVal = vehicleObj.smokeDueDate ? new Date(vehicleObj.smokeDueDate) : "";
    const data = sheet.getDataRange().getValues();
    
    for(let i=1; i<data.length; i++) {
      if(data[i][0] === vehicleObj.licensePlate) {
        sheet.getRange(i+1, 2, 1, 6).setValues([[vehicleObj.brandModel, vehicleObj.type, new Date(vehicleObj.taxDueDate), smokeDateVal, vehicleObj.imageUrl || "", vehicleObj.responsibleEmail || ""]]);
        SpreadsheetApp.flush();
        return {success: true, message: "อัปเดตข้อมูลรถสำเร็จ"};
      }
    }
    
    sheet.appendRow([vehicleObj.licensePlate, vehicleObj.brandModel, vehicleObj.type, new Date(vehicleObj.taxDueDate), smokeDateVal, vehicleObj.imageUrl || "", vehicleObj.responsibleEmail || ""]);
    SpreadsheetApp.flush();
    return {success: true, message: "เพิ่มรถคันใหม่สำเร็จ"};
  } catch(e) {
    return {success: false, message: e.toString()};
  }
}

function deleteVehicleAdmin(licensePlate, actorEmail) {
  try {
    writeAuditLog(actorEmail, "Vehicle Delete", `ลบข้อมูลรถ: ${licensePlate}`);
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Vehicles");
    const data = sheet.getDataRange().getValues();
    
    for(let i = data.length-1; i >= 1; i--) {
      if(data[i][0] === licensePlate) {
        sheet.deleteRow(i+1);
        SpreadsheetApp.flush();
        return {success: true, message: "ลบข้อมูลรถสำเร็จ"};
      }
    }
    return {success: false, message: "ไม่พบข้อมูลรถคันนี้"};
  } catch(e) {
    return {success: false, message: e.toString()};
  }
}

function getUserDropdownList() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    let res = [];
    for(let i=1; i<data.length; i++) {
      if(data[i][0]) {
        let name = data[i][1] || data[i][0];
        res.push({ title: `${name} (${data[i][3] || 'ไม่มีตำแหน่ง'})`, value: data[i][0] });
      }
    }
    return res;
  } catch(e) {
    return [];
  }
}

function checkVehicleDueDatesAlert() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Vehicles");
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    let today = new Date();
    today.setHours(0,0,0,0);
    
    let generalAlerts = []; 
    let directAlerts = {};
    
    for(let i=1; i<data.length; i++) {
      if(!data[i][0]) continue;
      
      let license = data[i][0];
      let taxDate = new Date(data[i][3]);
      let smokeDateStr = data[i][4];
      let respEmail = data[i][6] || "";
      
      if(isNaN(taxDate)) continue; // ข้ามถ้าวันที่ไม่ถูกต้อง
      
      let daysToTax = Math.floor((taxDate - today) / (1000 * 60 * 60 * 24));
      let msgs = [];
      
      if (daysToTax === 30 || daysToTax === 7 || daysToTax === 0 || daysToTax < 0) {
        let status = daysToTax < 0 ? `🚨 ขาดต่อมาแล้ว ${Math.abs(daysToTax)} วัน!` : (daysToTax === 0 ? '🔴 ครบกำหนดวันนี้!' : `🟡 เหลืออีก ${daysToTax} วัน`);
        msgs.push(`📄 <b>ภาษี:</b> ทะเบียน ${license} -> ${status}`);
      }
      
      if (smokeDateStr && String(smokeDateStr).trim() !== "") {
        let smokeDate = new Date(smokeDateStr);
        if(!isNaN(smokeDate)) {
          let daysToSmoke = Math.floor((smokeDate - today) / (1000 * 60 * 60 * 24));
          if (daysToSmoke === 15 || daysToSmoke === 3 || daysToSmoke === 0 || daysToSmoke < 0) {
            let status = daysToSmoke < 0 ? `🚨 เลยกำหนดมาแล้ว ${Math.abs(daysToSmoke)} วัน!` : (daysToSmoke === 0 ? '🔴 ครบกำหนดวันนี้!' : `🟡 เหลืออีก ${daysToSmoke} วัน`);
            msgs.push(`💨 <b>ควันดำ:</b> ทะเบียน ${license} -> ${status}`);
          }
        }
      }
      
      if (msgs.length > 0) {
        msgs.forEach(m => generalAlerts.push(m));
        if (respEmail) {
          if (!directAlerts[respEmail]) directAlerts[respEmail] = [];
          msgs.forEach(m => directAlerts[respEmail].push(m));
        }
      }
    }
    
    // แจ้งเตือนรายบุคคล (ผู้รับผิดชอบรถ)
    for(let email in directAlerts) {
      let tgId = getUserPersonalTgChatId(email);
      if(tgId) {
        let msg = `🚨 <b>แจ้งเตือนวาระยานพาหนะที่คุณดูแลรับผิดชอบอยู่</b>\n\n` + directAlerts[email].join("\n");
        sendTelegram(msg, "", "", tgId);
      }
    }
    
    // แจ้งเตือนเข้าระบบส่วนกลาง
    if (generalAlerts.length > 0) {
      let finalMsg = `🚨 <b>แจ้งเตือนยานพาหนะส่วนกลาง (ภาพรวมองค์กร)</b>\n\n` + generalAlerts.join("\n");
      sendTelegram(finalMsg); 
    }
  } catch (e) {
    console.error("checkVehicleDueDatesAlert Error: " + e);
  }
}

// =========================================
// 🚖 ระบบจองรถยนต์ส่วนกลาง (Vehicle Booking)
// =========================================
function checkVehicleAvailability(licensePlate, startDtStr, endDtStr) {
  try {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VehicleBookings");
    if(!sheet) return true;
    
    const reqStart = new Date(startDtStr).getTime(); 
    const reqEnd = new Date(endDtStr).getTime(); 
    const data = sheet.getDataRange().getValues();
    
    for(let i=1; i<data.length; i++) { 
      if(data[i][4] === licensePlate && data[i][10] !== "ไม่อนุมัติ" && data[i][10] !== "ยกเลิก") { 
        let bookStart = new Date(data[i][5]).getTime();
        let bookEnd = new Date(data[i][6]).getTime(); 
        if (reqStart < bookEnd && reqEnd > bookStart) { 
          return false; // เวลาชนกัน
        } 
      } 
    } 
    return true;
  } catch(e) { return true; }
}

function submitVehicleBooking(data) {
  try {
    writeAuditLog(data.email, "Vehicle Booking", `จองรถ: ${data.licensePlate}`); 
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("VehicleBookings");
    if(!sheet) { 
      sheet = ss.insertSheet("VehicleBookings");
      sheet.appendRow(["BookingID", "Timestamp", "Email", "FullName", "LicensePlate", "StartDateTime", "EndDateTime", "Destination", "Reason", "Passengers", "Status", "ApproverName", "CalendarEventId"]);
    }
    if(sheet.getRange(1, 13).getValue() !== "CalendarEventId") { sheet.getRange(1, 13).setValue("CalendarEventId"); }
    if(!checkVehicleAvailability(data.licensePlate, data.startDateTime, data.endDateTime)) { 
      return { success: false, message: "ขออภัย! มีผู้อื่นจองรถคันนี้ในช่วงเวลาดังกล่าวไปแล้ว" };
    }
    const bookingId = "VB" + new Date().getTime();
    const initialStatus = "รออนุมัติ (Pending)";
    sheet.appendRow([bookingId, new Date(), data.email, data.fullName, data.licensePlate, data.startDateTime, data.endDateTime, data.destination, data.reason, data.passengers, initialStatus, "", ""]);
    SpreadsheetApp.flush();
    
    let tgMsg = `🚗 <b>มีการขอจองรถส่วนกลางใหม่</b>\n👤 <b>ผู้จอง:</b> ${cleanTgText(data.fullName)}\n🚘 <b>ทะเบียนรถ:</b> ${cleanTgText(data.licensePlate)}\n📍 <b>สถานที่ไป:</b> ${cleanTgText(data.destination)}\n📅 <b>เวลา:</b> ${formatDateTimeSafe(data.startDateTime)} ถึง ${formatDateTimeSafe(data.endDateTime)}\n⏳ <b>สถานะ:</b> รอผู้อนุมัติรถส่วนกลาง`;
    sendTelegram(tgMsg, "", "", "", "vehicle"); // 🚀 บังคับให้ใช้บอทของรถส่วนกลาง
    
    return { success: true, message: "ส่งคำขอจองรถยนต์ส่วนกลางเรียบร้อยแล้ว" };
  } catch (e) { 
    return { success: false, message: "Error: " + e.toString() };
  }
}

function getMyVehicleBookings(email, userRoles, opt_nocache) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VehicleBookings"); 
  if (!sheet) return []; 
  const data = sheet.getDataRange().getValues();
  let results = [];
  
  for (let i = data.length - 1; i >= 1; i--) { 
    if (data[i][2] === email || userRoles.includes('Admin') || userRoles.includes('VehicleApprove')) { 
      results.push({ 
        rowId: i + 1, 
        bookingId: data[i][0], 
        timestamp: formatDateTimeSafe(data[i][1]), 
        email: data[i][2], 
        fullName: data[i][3], 
        licensePlate: data[i][4], 
        startDateTime: formatDateTimeSafe(data[i][5]), 
        endDateTime: formatDateTimeSafe(data[i][6]), 
        destination: data[i][7], 
        reason: data[i][8], 
        passengers: data[i][9], 
        status: data[i][10] 
      });
    } 
  } 
  return results;
}

function getPendingVehicleBookings(userRoles, opt_nocache) {
  if (!userRoles.some(r => ['Admin', 'VehicleApprove'].includes(r))) return [];
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VehicleBookings"); 
  if (!sheet) return []; 
  const data = sheet.getDataRange().getValues(); 
  let results = [];
  
  for (let i = data.length - 1; i >= 1; i--) {
    let status = data[i][10];
    // 🚀 เช็คให้แน่ใจว่าคนนั้นมีสิทธิ์ VehicleApprove ถึงจะเห็นคิว
    if (userRoles.includes('VehicleApprove') && status === "รออนุมัติ (Pending)") { 
      results.push({ 
        rowId: i + 1, 
        bookingId: data[i][0], 
        timestamp: formatDateTimeSafe(data[i][1]), 
        fullName: data[i][3], 
        licensePlate: data[i][4], 
        startDateTime: formatDateTimeSafe(data[i][5]), 
        endDateTime: formatDateTimeSafe(data[i][6]), 
        destination: data[i][7], 
        reason: data[i][8], 
        passengers: data[i][9], 
        status: data[i][10] 
      });
    }
  }
  return results;
}

function approveVehicleBooking(rowId, decision, approverName, email, comment = "") {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VehicleBookings");
    writeAuditLog(email, "Vehicle Approval", `พิจารณาจองรถแถว ${rowId} -> ${decision} (${comment})`);
    let nextStatus = decision === "ไม่อนุมัติ" ? "ไม่อนุมัติ" : "อนุมัติแล้ว (Approved)";
    sheet.getRange(rowId, 11).setValue(nextStatus);
    if(nextStatus === "อนุมัติแล้ว (Approved)") sheet.getRange(rowId, 12).setValue(approverName);
    
    if(nextStatus === "อนุมัติแล้ว (Approved)") {
      const data = sheet.getRange(rowId, 1, 1, 10).getValues()[0];
      const settings = getAllSettings(); 
      const calId = settings.vehicleCalendarId ? String(settings.vehicleCalendarId).trim() : String(settings.calendarId || "").trim();
      if(calId) { try { const cal = CalendarApp.getCalendarById(calId); if(cal) { const eventTitle = `[จองรถ] ${data[4]} - ${data[7]} (${data[3]})`; const ev = cal.createEvent(eventTitle, new Date(data[5]), new Date(data[6]), {description: data[8]}); sheet.getRange(rowId, 13).setValue(ev.getId()); } } catch(e) {} }
    }
    SpreadsheetApp.flush();
    const licensePlate = sheet.getRange(rowId, 5).getValue(); 
    const reqName = sheet.getRange(rowId, 4).getValue(); 
    const reqEmail = sheet.getRange(rowId, 3).getValue();
    
    let tgMsg = `🔔 <b>อัปเดตสถานะจองรถ</b>\n👤 ผู้จอง: ${cleanTgText(reqName)}\n🚘 รถ: ${cleanTgText(licensePlate)}\n✅ <b>สถานะล่าสุด: ${cleanTgText(nextStatus)}</b>\n💬 <b>ความเห็น:</b> ${cleanTgText(comment || "-")}`;
    let tgId = getUserPersonalTgChatId(reqEmail); 
    if(tgId) sendTelegram(tgMsg, "", "", tgId, "vehicle"); // 🚀 บังคับให้ใช้บอทของรถส่วนกลาง
    sendTelegram(tgMsg, "", "", "", "vehicle"); // 🚀 บังคับให้ใช้บอทของรถส่วนกลาง
    
    return { success: true, message: "บันทึกผลสำเร็จ" };
  } catch(e) { return { success: false, message: e.toString() }; }
}


function getVehicleTimelineBookings(licensePlate, dateStr) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VehicleBookings");
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    const results = [];
    for(let i=1; i<data.length; i++) {
      if(data[i][4] === licensePlate && data[i][10] !== "ยกเลิก" && data[i][10] !== "ไม่อนุมัติ") {
        let sDate = new Date(data[i][5]);
        let eDate = new Date(data[i][6]);
        let sStr = Utilities.formatDate(sDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
        let eStr = Utilities.formatDate(eDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
        
        if(sStr === dateStr || eStr === dateStr || (new Date(dateStr) >= sDate && new Date(dateStr) <= eDate)) {
          results.push({
            start: Utilities.formatDate(sDate, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss"),
            end: Utilities.formatDate(eDate, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss"),
            title: data[i][7], fullName: data[i][3], status: data[i][10]
          });
        }
      }
    }
    return results;
  } catch(e) { return []; }
}

// =========================================
// 🕒 ดึงข้อมูลสำหรับวาด Timeline ห้องประชุม
// =========================================
function getTimelineBookings(roomName, dateStr) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("RoomBookings");
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    const results = [];
    for(let i=1; i<data.length; i++) {
      if(data[i][4] === roomName && data[i][10] !== "ยกเลิก" && data[i][10] !== "ไม่อนุมัติ") {
        let sDate = new Date(data[i][5]);
        let eDate = new Date(data[i][6]);
        let sStr = Utilities.formatDate(sDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
        let eStr = Utilities.formatDate(eDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
        
        // เช็คว่าวันนั้นอยู่ในช่วงที่จองไว้หรือไม่
        if(sStr === dateStr || eStr === dateStr || (new Date(dateStr) >= sDate && new Date(dateStr) <= eDate)) {
          results.push({
            start: Utilities.formatDate(sDate, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss"),
            end: Utilities.formatDate(eDate, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss"),
            title: data[i][7], fullName: data[i][3], status: data[i][10]
          });
        }
      }
    }
    return results;
  } catch(e) { return []; }
}
// =========================================
// 🏆 ดึงข้อมูล 5 อันดับ พนักงานที่ลาน้อยที่สุด
// =========================================
function getTopAttendanceStats() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const userSheet = ss.getSheetByName("Users");
    const leaveSheet = ss.getSheetByName("LeaveRequests");

    let users = {};
    if(userSheet) {
      const uData = userSheet.getDataRange().getValues();
      if (uData.length === 0) return [];
      let headers = uData[0];
      let emailIdx = headers.indexOf("Email") > -1 ? headers.indexOf("Email") : 0;
      let nameIdx = headers.indexOf("FullName") > -1 ? headers.indexOf("FullName") : 1;
      let avatarIdx = headers.indexOf("Avatar") > -1 ? headers.indexOf("Avatar") : 7;
      let roleIdx = headers.indexOf("Role") > -1 ? headers.indexOf("Role") : 8;
      
      for(let i=1; i<uData.length; i++) {
        let email = uData[i][emailIdx];
        let role = uData[i][roleIdx];
        let fullName = uData[i][nameIdx] || '';
        
        if(email && !String(role).includes("Admin") && fullName.trim() !== "") { 
          users[email] = {
            email: email,
            fullName: fullName,
            avatar: uData[i][avatarIdx] || 'https://cdn-icons-png.flaticon.com/512/3135/3135715.png',
            totalLeaveDays: 0
          };
        }
      }
    }

    if(leaveSheet) {
      const lData = leaveSheet.getDataRange().getValues();
      if (lData.length > 0) {
        let headers = lData[0];
        let emailIdx = headers.indexOf("Email") > -1 ? headers.indexOf("Email") : 1;
        let statusIdx = headers.indexOf("Status") > -1 ? headers.indexOf("Status") : 9; 
        let daysIdx = headers.indexOf("WorkingDays") > -1 ? headers.indexOf("WorkingDays") : 6;
        if(emailIdx === -1) emailIdx = 1;
        if(daysIdx === -1) daysIdx = 6;
        if(statusIdx === -1) statusIdx = 9;
        
        for(let i=1; i<lData.length; i++) {
          let email = lData[i][emailIdx];
          let status = String(lData[i][statusIdx] || "");
          let days = parseFloat(lData[i][daysIdx]) || 0;
          if (status.includes("อนุมัติเสร็จสิ้น") || status.includes("อนุมัติแล้ว")) {
            if(users[email]) {
              users[email].totalLeaveDays += days;
            }
          }
        }
      }
    }

    let sortedUsers = Object.values(users).sort((a, b) => a.totalLeaveDays - b.totalLeaveDays);
    return sortedUsers.slice(0, 5);
  } catch(e) {
    console.error("getTopAttendanceStats Error: " + e);
    return [];
  }
}

// ==========================================
// 🛡️ ระบบตั้งค่าผู้ใช้โดยแอดมิน (Admin User Management)
// ==========================================
function getAllUsersAdmin(opt_nocache) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues(); 
    let users = [];
    for (let i = 1; i < data.length; i++) { 
      let dStart = data[i][11] ? Utilities.formatDate(new Date(data[i][11]), Session.getScriptTimeZone(), "yyyy-MM-dd") : "";
      let dEnd = data[i][12] ? Utilities.formatDate(new Date(data[i][12]), Session.getScriptTimeZone(), "yyyy-MM-dd") : "";
      let sDate = data[i][14] ? Utilities.formatDate(new Date(data[i][14]), Session.getScriptTimeZone(), "yyyy-MM-dd") : "";
      
      users.push({ 
        email: data[i][0] || "", 
        fullName: data[i][1] || "", 
        phone: data[i][2] || "", 
        position: data[i][3] || "", 
        affiliation: data[i][4] || "",
        address: data[i][5] || "",
        signature: data[i][6] || "",
        avatar: data[i][7] || "", 
        roles: String(data[i][8] || 'User').split(',').map(r => r.trim()), 
        accumulatedLeave: parseFloat(data[i][9]) || 0, 
        delegateEmail: data[i][10] || "", 
        delegateStart: dStart, 
        delegateEnd: dEnd,
        startDate: sDate, 
        employmentType: data[i][15] || "" 
      });
    } 
    return users;
  } catch(e) {
    return [];
  }
}

function saveUserRoleAdmin(userObj, actorEmail) {
  try {
    writeAuditLog(actorEmail, "Change Role", `อัปเดตข้อมูลพนักงาน: ${userObj.email}`);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) { 
      if (data[i][0] === userObj.email) {
        sheet.getRange(i + 1, 2).setValue(userObj.fullName || "");
        sheet.getRange(i + 1, 4).setValue(userObj.position || "");
        sheet.getRange(i + 1, 9).setValue(Array.isArray(userObj.roles) ? userObj.roles.join(',') : 'User');
        sheet.getRange(i + 1, 10).setValue(userObj.accumulatedLeave || 0);
        sheet.getRange(i + 1, 11).setValue(userObj.delegateEmail || ""); 
        sheet.getRange(i + 1, 12).setValue(userObj.delegateStart || "");
        sheet.getRange(i + 1, 13).setValue(userObj.delegateEnd || "");
        sheet.getRange(i + 1, 15).setValue(userObj.startDate || ""); 
        sheet.getRange(i + 1, 16).setValue(userObj.employmentType || ""); 
        SpreadsheetApp.flush(); 
        return {success: true, message: "อัปเดตข้อมูลผู้ใช้งานสำเร็จ"};
      } 
    } 
    return {success: false, message: "ไม่พบผู้ใช้งาน"};
  } catch(e) { return {success: false, message: e.toString()}; }
}

function addUserAdmin(userObj, actorEmail) {
  try {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === userObj.email) return {success: false, message: "อีเมลนี้มีอยู่ในระบบแล้ว"};
    }
    let rolesStr = Array.isArray(userObj.roles) ? userObj.roles.join(',') : 'User';
    sheet.appendRow([
      userObj.email, userObj.fullName || "", "", userObj.position || "", "", "", "", "",
      rolesStr, userObj.accumulatedLeave || 0, "", "", "", "", userObj.startDate || "", userObj.employmentType || ""
    ]);
    writeAuditLog(actorEmail, "Add User", `เพิ่มพนักงานใหม่: ${userObj.email}`);
    SpreadsheetApp.flush();
    return {success: true, message: "เพิ่มพนักงานใหม่สำเร็จ"};
  } catch(e) { return {success: false, message: e.toString()}; }
}

function deleteUserAdmin(email, actorEmail) {
  try {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
    const data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][0] === email) {
        sheet.deleteRow(i + 1);
        writeAuditLog(actorEmail, "Delete User", `ลบพนักงาน: ${email}`);
        SpreadsheetApp.flush();
        return {success: true, message: "ลบข้อมูลพนักงานสำเร็จ"};
      }
    }
    return {success: false, message: "ไม่พบอีเมลในระบบ"};
  } catch(e) { return {success: false, message: e.toString()}; }
}

// ==========================================
// 📚 ระบบหนังสือเวียนภายใน (E-Document)
// ==========================================
function getEDocuments(userEmail, userRoles) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("EDocuments");
    if (!sheet) return { inbox: [], manage: [] };
    
    const data = sheet.getDataRange().getValues();
    let inbox = [];
    let manage = [];
    let isAdminOrHR = userRoles.includes('Admin') || userRoles.includes('HR') || userRoles.includes('Approve1');
    
    for (let i = data.length - 1; i >= 1; i--) {
      let docId = data[i][0];
      let timestamp = formatDateTimeSafe(data[i][1]);
      let senderName = data[i][3];
      let docNumber = data[i][4];
      let title = data[i][5];
      let desc = data[i][6];
      let fileUrl = data[i][7];
      let targetsStr = data[i][8] || "[]";
      let acksStr = data[i][9] || "[]";
      
      let targets = [];
      let acks = [];
      try { targets = JSON.parse(targetsStr); } catch(e){}
      try { acks = JSON.parse(acksStr); } catch(e){}
      
      let isTarget = targets.includes("ทั้งหมด") || targets.includes(userEmail);
      let hasAcked = acks.some(a => a.email === userEmail);
      
      if (isTarget) {
        inbox.push({
          rowId: i + 1, docId, timestamp, senderName, docNumber, title, desc, fileUrl, 
          isAcked: hasAcked,
          ackTime: hasAcked ? formatDateTimeSafe(acks.find(a => a.email === userEmail).time) : "-"
        });
      }
      
      if (isAdminOrHR) {
        manage.push({
          rowId: i + 1, docId, timestamp, senderName, docNumber, title, desc, fileUrl,
          targets: targets,
          acks: acks,
          totalTargets: targets.includes("ทั้งหมด") ? "ทุกคน" : targets.length,
          totalAcks: acks.length
        });
      }
    }
    return { inbox: inbox, manage: manage };
  } catch (e) {
    return { inbox: [], manage: [] };
  }
}

function submitEDocument(data) {
  try {
    writeAuditLog(data.senderEmail, "Send E-Doc", `ส่งหนังสือเวียน: ${data.title}`);
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("EDocuments");
    
    if(!sheet) {
      sheet = ss.insertSheet("EDocuments");
      sheet.appendRow(["DocID", "Timestamp", "SenderEmail", "SenderName", "DocNumber", "Title", "Description", "FileUrl", "Targets", "AckData"]);
    }
    
    let fileUrl = "ไม่มีไฟล์แนบ";
    if (data.fileBase64 && data.fileName) {
      const settings = getAllSettings();
      if(settings.folderId) {
        const folderId = cleanFolderId(settings.folderId);
        const folder = DriveApp.getFolderById(folderId);
        const contentType = data.fileBase64.substring(5, data.fileBase64.indexOf(';'));
        const bytes = Utilities.base64Decode(data.fileBase64.split(',')[1]);
        
        let newFile = folder.createFile(Utilities.newBlob(bytes, contentType, "เวียน_" + data.fileName));
        newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        fileUrl = newFile.getUrl();
      }
    }
    
    const docId = "DOC" + new Date().getTime();
    let targetsJson = JSON.stringify(data.targets || []);
    
    sheet.appendRow([docId, new Date(), data.senderEmail, data.senderName, data.docNumber, data.title, data.description, fileUrl, targetsJson, "[]"]);
    SpreadsheetApp.flush();
    
    let tgMsg = `📢 <b>มีหนังสือเวียนใหม่!</b>\nเรื่อง: ${cleanTgText(data.title)}\nเลขที่: ${cleanTgText(data.docNumber)}\nจาก: ${cleanTgText(data.senderName)}\n\n👉 กรุณาเข้าสู่ระบบเพื่อเปิดอ่านและกด <b>"รับทราบ"</b>`;
    if(data.targets.includes("ทั้งหมด")) {
      sendTelegram(tgMsg, fileUrl);
    } else {
      data.targets.forEach(email => {
        let tgId = getUserPersonalTgChatId(email);
        if(tgId) sendTelegram(tgMsg, fileUrl, "", tgId);
      });
    }
    
    return { success: true, message: "ส่งหนังสือเวียนเรียบร้อยแล้ว" };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function acknowledgeEDocument(rowId, userEmail, userName) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("EDocuments");
    let acksStr = sheet.getRange(rowId, 10).getValue() || "[]";
    let acks = [];
    try { acks = JSON.parse(acksStr); } catch(e){}
    
    if(!acks.some(a => a.email === userEmail)) {
      acks.push({ email: userEmail, name: userName, time: new Date() });
      sheet.getRange(rowId, 10).setValue(JSON.stringify(acks));
      SpreadsheetApp.flush();
    }
    return { success: true, message: "บันทึกการรับทราบเรียบร้อยแล้ว" };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function deleteEDocument(rowId, userEmail) {
  try {
    writeAuditLog(userEmail, "Delete E-Doc", `ลบหนังสือเวียนแถวที่ ${rowId}`);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("EDocuments");
    sheet.deleteRow(rowId);
    SpreadsheetApp.flush();
    return { success: true, message: "ลบหนังสือเวียนเรียบร้อยแล้ว" };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// ==========================================
// 📂 ระบบคลังเอกสาร (Documents)
// ==========================================
function getDocuments() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Documents");
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    let docs = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][1]) { 
        docs.push({
          category: data[i][0] || "ทั่วไป",
          name: data[i][1] || "",
          updatedDate: data[i][2] || "",
          link: data[i][3] || "#"
        });
      }
    }
    return docs;
  } catch(e) { return []; }
}

function addDocument(payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Documents");
    if (!sheet) {
      sheet = ss.insertSheet("Documents");
      sheet.appendRow(["หมวดหมู่", "ชื่อแบบฟอร์ม", "วันที่อัปเดต", "ลิงก์ดาวน์โหลด"]);
    }
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
    sheet.appendRow([payload.category, payload.name, today, payload.link]);
    SpreadsheetApp.flush();
    return { success: true, message: "อัปโหลดเอกสารลงคลังเรียบร้อยแล้ว" };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function getStaffDirectory() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    let staffList = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) { 
        staffList.push({
          email: data[i][0] || "",
          fullName: data[i][1] || "ไม่มีชื่อ",
          phone: data[i][2] || "-",
          position: data[i][3] || "-",
          affiliation: data[i][4] || "-",
          avatar: data[i][7] || "https://ui-avatars.com/api/?name=" + encodeURIComponent(data[i][1] || "U") + "&background=random&color=fff",
        });
      }
    }
    staffList.sort((a, b) => a.fullName.localeCompare(b.fullName, 'th'));
    return staffList;
  } catch(e) { return []; }
}

// ==========================================
// 🔑 ฟังก์ชันเปลี่ยนรหัสผ่าน (เวอร์ชันแก้ทางค้าง)
// ==========================================
function changeUserPassword(email, oldPassword, newPassword) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Users");
    if (!sheet) return { success: false, message: "ไม่พบชีท Users กรุณาตรวจสอบชื่อชีท" };

    const data = sheet.getDataRange().getDisplayValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === email) {
        const currentPassword = data[i][16] ? data[i][16].toString().trim() : "";
        const inputOld = oldPassword ? oldPassword.toString().trim() : "";
        if (currentPassword !== "" && currentPassword !== inputOld) {
          return { success: false, message: "รหัสผ่านเดิมไม่ถูกต้อง" };
        }
        sheet.getRange(i + 1, 17).setValue(newPassword);
        SpreadsheetApp.flush(); 
        return { success: true, message: "เปลี่ยนรหัสผ่านสำเร็จแล้ว!" };
      }
    }
    return { success: false, message: "ไม่พบอีเมลผู้ใช้ในระบบ" };
  } catch (e) {
    return { success: false, message: "เกิดข้อผิดพลาด: " + e.toString() };
  }
}

// ==========================================
// ⚙️ ระบบดึงและบันทึกการตั้งค่า (รองรับ Key เดิมที่มีอยู่)
// ==========================================
function getAppConfig() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
    let config = {};
    if (!sheet) return { all: config, flowRules: [], telegram: { leave: { token: "", chatId: "" }, vehicle: { token: "", chatId: "" }, room: { token: "", chatId: "" } } };

    const data = sheet.getDataRange().getValues();
    data.forEach(row => {
      if (row[0]) config[row[0]] = row[1];
    });

    // 🚀 เซ็ตค่าเริ่มต้นให้ครบทั้ง 3 ระบบ
    let tgData = { leave: { token: "", chatId: "" }, vehicle: { token: "", chatId: "" }, room: { token: "", chatId: "" } };
    
    if (config["TELEGRAM"]) {
       try {
         let parsedTg = JSON.parse(config["TELEGRAM"]);
         if (parsedTg.leave) tgData.leave = { ...tgData.leave, ...parsedTg.leave };
         if (parsedTg.vehicle) tgData.vehicle = { ...tgData.vehicle, ...parsedTg.vehicle };
         if (parsedTg.room) tgData.room = { ...tgData.room, ...parsedTg.room };
       } catch(e){}
    } else {
       // รองรับระบบเก่า
       tgData.leave.token = config["telegramToken"] || "";
       tgData.leave.chatId = config["telegramChatId"] || "";
    }

    return {
      all: config, 
      flowRules: config["FLOW_RULES"] ? JSON.parse(config["FLOW_RULES"]) : [],
      telegram: tgData
    };
  } catch (e) {
    console.error("getAppConfig Error: " + e);
    return { all: {}, flowRules: [], telegram: { leave: { token: "", chatId: "" }, vehicle: { token: "", chatId: "" }, room: { token: "", chatId: "" } } };
  }
}

function saveAppConfig(flowRules, telegramData) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
    if (!sheet) return { success: false, message: "ไม่พบชีท Settings" };

    const data = sheet.getDataRange().getValues();
    const updateOrAppend = (key, value) => {
      let rowIndex = -1;
      for (let i = 0; i < data.length; i++) {
        if (data[i][0] === key) { rowIndex = i + 1; break; }
      }
      if (rowIndex > -1) {
        sheet.getRange(rowIndex, 2).setValue(value);
      } else {
        sheet.appendRow([key, value]);
      }
    };

    updateOrAppend("FLOW_RULES", JSON.stringify(flowRules));
    updateOrAppend("TELEGRAM", JSON.stringify(telegramData));

    SpreadsheetApp.flush();
    return { success: true, message: "บันทึกการตั้งค่าเรียบร้อยแล้ว" };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}
// ==========================================
// 🗂️ ระบบแฟ้มทะเบียนประวัติ ก.พ. 7 (KP7 Records)
// ==========================================
function getKp7Record(email) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("KP7Records");
    if (!sheet) return null;
    
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === email) {
        let trainings = [];
        try { trainings = JSON.parse(data[i][8] || "[]"); } catch(e){}
        return {
          email: data[i][0],
          idCard: data[i][1] ? String(data[i][1]) : "",
          birthDate: data[i][2] ? formatDateForInput(data[i][2]) : "",
          bloodGroup: data[i][3] || "",
          religion: data[i][4] || "",
          salary: data[i][5] || "",
          bankName: data[i][6] || "",
          bankAccount: data[i][7] ? String(data[i][7]) : "",
          trainings: trainings
        };
      }
    }
    return null;
  } catch(e) {
    return null;
  }
}

function saveKp7Record(payload, actorEmail) {
  try {
    writeAuditLog(actorEmail, "Update KP7", `อัปเดตข้อมูล ก.พ. 7 ของ ${payload.email}`);
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("KP7Records");
    
    if (!sheet) {
      sheet = ss.insertSheet("KP7Records");
      sheet.appendRow(["Email", "ID_Card", "BirthDate", "BloodGroup", "Religion", "Salary", "BankName", "BankAccount", "TrainingHistory"]);
    }
    
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === payload.email) {
        rowIndex = i + 1;
        break;
      }
    }
    
    let bd = payload.birthDate ? new Date(payload.birthDate) : "";
    let trainingsStr = JSON.stringify(payload.trainings || []);
    
    const rowData = [
      payload.email, 
      "'" + payload.idCard, // ใส่ ' ดักไว้ไม่ให้เลข 13 หลักกลายเป็นเลขวิทยาศาสตร์
      bd, 
      payload.bloodGroup, 
      payload.religion, 
      payload.salary, 
      payload.bankName, 
      "'" + payload.bankAccount, 
      trainingsStr
    ];
    
    if (rowIndex > -1) {
      sheet.getRange(rowIndex, 1, 1, 9).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }
    
    SpreadsheetApp.flush();
    return { success: true, message: "บันทึกข้อมูลทะเบียนประวัติ ก.พ. 7 เรียบร้อยแล้ว" };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function saveAllSettings(settingsObj, actorEmail) {
  try {
    writeAuditLog(actorEmail, "System Config", "แอดมินแก้ไขการตั้งค่าระบบ"); 
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
    if (!sheet) { 
      sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Settings"); 
      sheet.appendRow(["Key", "Value"]);
    }
    
    const data = sheet.getDataRange().getValues();
    for (let key in settingsObj) { 
      let found = false;
      for (let i = 1; i < data.length; i++) { 
        if (data[i][0] === key) { 
          sheet.getRange(i + 1, 2).setValue(settingsObj[key]);
          found = true; 
          break; 
        } 
      } 
      if (!found) {
        sheet.appendRow([key, settingsObj[key]]);
      }
    }
    
    SpreadsheetApp.flush(); 
    return { success: true, message: "บันทึกการตั้งค่าระบบเรียบร้อยแล้ว" };
  } catch (e) { 
    return { success: false, message: e.toString() };
  }
}
