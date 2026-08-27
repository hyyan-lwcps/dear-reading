function doGet(e) {
  var template = HtmlService.createTemplateFromFile('Index');
  template.currentFloor = e.parameter.floor || "";
  return template.evaluate()
      .setTitle("2026-2027 D.E.A.R. 午讀評分系統")
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 🏢 根據樓層傳回班別對照表
function getClassesForFloor(floor) {
  var floorMap = {
    "1": ["1A", "1B", "1C", "1D"],
    "2": ["2A", "2B", "2C", "2D"],
    "3": ["3A", "3B", "3C", "3D"],
    "4": ["4A", "4B", "4C", "4D"],
    "5": ["5A", "5B", "5C", "5D"],
    "6": ["6A", "6B", "6C", "6D"]
  };
  return floorMap[floor] || [];
}

// 📝 提交午讀數據 (完全適應：時間戳記, 月份, 班別, 基本達標, 國安圖書, 主題圖書)
function submitReadData(dataList, customDateStr) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rows = [];
  
  // 確定評分日期與時間戳記
  var now = new Date();
  var todayStr = customDateStr ? customDateStr : Utilities.formatDate(now, "GMT+8", "yyyy-MM-dd");
  var nowTimestamp = Utilities.formatDate(now, "GMT+8", "yyyy-MM-dd HH:mm:ss");
  
  // 自動計算月份 (例如 "2026-09-01" 擷取出的月份數字)
  var targetDateObj = new Date(todayStr.replace(/-/g, "/"));
  var monthNum = targetDateObj.getMonth() + 1;

  // 防重複提交檢查 (檢查 A 欄日期與 C 欄班別)
  var existingData = sheet.getDataRange().getValues();
  var submittedClasses = [];

  for (var i = 1; i < existingData.length; i++) {
    var rowDate = existingData[i][0]; // A 欄：時間戳記或日期
    var rowClass = existingData[i][2]; // C 欄：班別

    if (rowDate instanceof Date) {
      rowDate = Utilities.formatDate(rowDate, "GMT+8", "yyyy-MM-dd");
    } else if (typeof rowDate === 'string' && rowDate.length >= 10) {
      rowDate = rowDate.substring(0, 10);
    }
    
    if (rowDate === todayStr) {
      submittedClasses.push(rowClass); 
    }
  }

  var duplicates = [];
  dataList.forEach(function(item) {
    if (submittedClasses.indexOf(item.className) !== -1) {
      duplicates.push(item.className);
    }
  });

  if (duplicates.length > 0) {
    return { status: 'duplicate', duplicateClasses: duplicates, targetDate: todayStr };
  }

  // 按照原排版寫入：[A:時間戳記, B:月份, C:班別, D:基本達標, E:國安圖書, F:主題圖書]
  dataList.forEach(function(item) {
    rows.push([
      nowTimestamp,                   // A欄：時間戳記 (或當日日期 todayStr，視習慣)
      monthNum,                       // B欄：月份
      item.className,                 // C欄：班別
      item.basic ? "YES" : "NO",      // D欄：基本達標
      item.ns ? "YES" : "NO",         // E欄：國安圖書
      item.theme ? "YES" : "NO"       // F欄：主題圖書
    ]);
  });

  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }

  return { status: 'success' };
}

// 📊 取得指定月份的龍虎榜統計資料
function getLeaderboardData(targetMonth) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var stats = {};

  for (var i = 1; i < data.length; i++) {
    var dateVal = data[i][0]; // A 欄：時間戳記/日期
    var recordMonth = data[i][1]; // B 欄：月份
    var cls = data[i][2];     // C 欄：班別
    var basic = data[i][3];   // D 欄：基本達標
    var ns = data[i][4];      // E 欄：國安圖書
    var theme = data[i][5];   // F 欄：主題圖書

    // 若 B 欄無數字，嘗試從 A 欄日期推算月份
    if (!recordMonth && dateVal) {
      if (dateVal instanceof Date) {
        recordMonth = dateVal.getMonth() + 1;
      } else if (typeof dateVal === 'string') {
        var parts = dateVal.split('-');
        if (parts.length >= 2) recordMonth = parseInt(parts[1], 10);
      }
    }

    if (recordMonth == targetMonth) {
      if (!stats[cls]) {
        stats[cls] = { basic: 0, ns: 0, theme: 0, total: 0 };
      }
      if (basic === "YES") stats[cls].basic += 1;
      if (ns === "YES") stats[cls].ns += 1;
      if (theme === "YES") stats[cls].theme += 1;
      
      stats[cls].total = stats[cls].basic + stats[cls].ns + stats[cls].theme;
    }
  }

  var leaderboard = [];
  Object.keys(stats).forEach(function(cls) {
    leaderboard.push({
      className: cls,
      basic: stats[cls].basic,
      ns: stats[cls].ns,
      theme: stats[cls].theme,
      total: stats[cls].total
    });
  });

  return leaderboard;
}

// 📜 取得管理員後台歷史紀錄明細
function getAdminData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var logs = [];

  for (var i = data.length - 1; i >= 1; i--) {
    var dateStr = data[i][0]; // A 欄：時間戳記
    if (dateStr instanceof Date) {
      dateStr = Utilities.formatDate(dateStr, "GMT+8", "yyyy-MM-dd HH:mm");
    }
    
    logs.push({
      rowIndex: i + 1,
      timestamp: dateStr,
      month: data[i][1],      // B 欄：月份
      className: data[i][2],  // C 欄：班別
      basic: data[i][3],      // D 欄：基本達標
      ns: data[i][4],         // E 欄：國安圖書
      theme: data[i][5]       // F 欄：主題圖書
    });
  }
  return logs;
}

// ✏️ 後台更新整列評分 (更新 D、E、F 欄)
function updateFullRecord(rowIndex, basic, ns, theme) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // 第 4 欄 (D 欄) 開始，更新 3 個欄位 (D:基本, E:國安, F:主題)
  sheet.getRange(rowIndex, 4, 1, 3).setValues([[basic, ns, theme]]);
  return { status: 'success' };
}

// 🗑️ 刪除多筆指定列
function deleteMultipleRecords(rowIndices) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  rowIndices.sort(function(a, b) { return b - a; });

  rowIndices.forEach(function(r) {
    sheet.deleteRow(r);
  });

  return { status: 'success', count: rowIndices.length };
}
