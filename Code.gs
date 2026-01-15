// Code.gs
// 員工考勤系統後端邏輯

/**
 * 處理 HTTP GET 請求，回傳 Index.html 頁面
 */
function doGet(e) {
  var template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate()
    .setTitle('員工考勤系統')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 取得連結的試算表 - [Attendance] 分頁
 */
function getAttendanceSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Attendance");
  
  if (!sheet) {
    sheet = ss.insertSheet("Attendance");
    sheet.appendRow(["員工編號", "姓名", "日期", "簽到時間", "簽退時間", "備註"]);
  } else {
    // 檢查是否有員工編號欄位 (相容性考慮)
    var header = sheet.getRange(1, 1, 1, 1).getValue();
    if (header !== "員工編號") {
       sheet.insertColumnBefore(1);
       sheet.getRange(1, 1).setValue("員工編號");
    }
  }
  return sheet;
}

/**
 * 取得或建立 [Users] 分頁
 * 預設會建立一筆測試資料
 */
function getUsersSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Users");
  
  if (!sheet) {
    sheet = ss.insertSheet("Users");
    sheet.appendRow(["員工編號", "密碼", "姓名"]);
    // 預設測試資料 (E001 為管理員)
    sheet.appendRow(["E001", "123456", "系統管理員"]);
    sheet.appendRow(["A001", "password", "王小明"]);
  }
  return sheet;
}

/**
 * 登入驗證
 * @param {string} id - 員工編號
 * @param {string} password - 密碼
 */
function checkLogin(id, password) {
  var sheet = getUsersSheet();
  var data = sheet.getDataRange().getValues();
  
  var inputId = id ? id.toString().trim().toUpperCase() : "";
  var inputPwd = password ? password.toString().trim() : "";

  for (var i = 1; i < data.length; i++) {
    var sheetId = data[i][0] ? data[i][0].toString().trim().toUpperCase() : "";
    var sheetPwd = data[i][1] ? data[i][1].toString().trim() : "";
    
    if (sheetId === inputId && sheetPwd === inputPwd) {
      return { status: 'success', name: data[i][2] };
    }
  }
  
  console.log("登入失敗詳情 - 輸入 ID: [" + inputId + "], 查無匹配。");
  return { status: 'error', message: '編號或密碼錯誤' };
}

/**
 * 根據 ID 獲取姓名
 */
function getNameById(id) {
  var sheet = getUsersSheet();
  var data = sheet.getDataRange().getValues();
  var targetId = id ? id.toString().trim().toUpperCase() : "";
  
  for (var i = 1; i < data.length; i++) {
    var sheetId = data[i][0] ? data[i][0].toString().trim().toUpperCase() : "";
    if (sheetId === targetId) {
      return data[i][2];
    }
  }
  return null;
}

/**
 * 處理打卡請求
 * @param {Object} data - {id: string, action: 'checkin'|'checkout', note: string}
 */
function processPunch(data) {
  var id = data.id.trim();
  var action = data.action;
  var note = data.note || "";
  
  // 安全檢查：重新抓取姓名
  var name = getNameById(id);
  if (!name) {
    return {status: 'error', message: '無效的員工編號'};
  }

  var sheet = getAttendanceSheet();
  var now = new Date();
  var dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy/MM/dd");
  var timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "HH:mm:ss");

  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();
  var rowIndex = -1;

  var targetId = id.trim().toUpperCase();
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var rowId = row[0] ? row[0].toString().trim().toUpperCase() : "";
    var rowDate = row[2];
    var rowDateStr = (rowDate instanceof Date) ? 
      Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "yyyy/MM/dd") : rowDate;
      
    if (rowId === targetId && rowDateStr == dateStr) {
      rowIndex = i + 1;
      break;
    }
  }

  try {
    if (action === 'checkin') {
      if (rowIndex !== -1) {
        return {status: 'error', message: '您今天已經簽到過了！'};
      }
      // [員工編號, 姓名, 日期, 簽到時間, 簽退時間, 備註]
      sheet.appendRow([targetId, name, dateStr, timeStr, "", note]);

      var displayDate = Utilities.formatDate(now, Session.getScriptTimeZone(), "MM月dd日");
      var displayTime = Utilities.formatDate(now, Session.getScriptTimeZone(), "HH點mm分ss秒");

      return {
        status: 'success', 
        message: '簽到成功！',
        data: { id: targetId, name: name, date: displayDate, inTime: displayTime, outTime: "", note: note },
        action: "checkin"
      };
      
    } else if (action === 'checkout') {
      if (rowIndex === -1) {
        return {status: 'error', message: '找不到今日簽到紀錄，無法簽退！'};
      }
      // 更新簽退時間 (第 5 欄)
      sheet.getRange(rowIndex, 5).setValue(timeStr);
      
      var oldNote = values[rowIndex - 1][5];
      var newNote = oldNote ? (oldNote + " | [退]" + note) : ("[退]" + note);
      if(note) { 
          sheet.getRange(rowIndex, 6).setValue(newNote);
      }

      var inTimeRaw = values[rowIndex - 1][3];
      var displayInTime = (inTimeRaw instanceof Date) ? 
          Utilities.formatDate(inTimeRaw, Session.getScriptTimeZone(), "HH點mm分ss秒") : String(inTimeRaw);
      
      var displayDate = Utilities.formatDate(now, Session.getScriptTimeZone(), "MM月dd日");
      var displayTime = Utilities.formatDate(now, Session.getScriptTimeZone(), "HH點mm分ss秒");
      
      return {
        status: 'success', 
        message: '簽退成功！',
        data: { id: targetId, name: name, date: displayDate, inTime: displayInTime, outTime: displayTime, note: newNote },
        action: "checkout"
      };
    }
  } catch (e) {
    return {status: 'error', message: '系統錯誤: ' + e.toString()};
  }
}


/**
 * 取得出勤紀錄 (基於 ID 進行過濾)
 * @param {string} filterId - 發送請求的員工 ID
 */
function getAttendanceData(filterId) {
  var sheet = getAttendanceSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }
  
  var data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  var targetId = filterId ? filterId.toString().trim().toUpperCase() : "";
  
  // 設定管理員名單 (可看全體紀錄)
  var adminList = ["E000", "A001"];
  var isAdmin = adminList.includes(targetId);
  
  var filteredData = data.filter(function(row) {
    var rowId = row[0] ? row[0].toString().trim().toUpperCase() : "";
    return isAdmin || rowId === targetId;
  });
  
  var formattedData = filteredData.map(function(row) {
    var dateVal = row[2];
    var dateStr = (dateVal instanceof Date) ? 
      Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "MM月dd日") : String(dateVal);
      
    var inTime = (row[3] instanceof Date) ? 
      Utilities.formatDate(row[3], Session.getScriptTimeZone(), "HH點mm分ss秒") : String(row[3]);
      
    var outTime = (row[4] instanceof Date) ? 
      Utilities.formatDate(row[4], Session.getScriptTimeZone(), "HH點mm分ss秒") : String(row[4]);
      
    var note = row[5] ? String(row[5]) : "";
    
    return [
      String(row[0]), // 編號
      String(row[1]), // 姓名
      dateStr,
      inTime,
      outTime,
      note
    ];
  });
  
  return formattedData;
}


