// ==========================================
// DEAR 午讀評分系統 - 後端核心邏輯 (Code.gs)
// ==========================================

// 1. Web App 入口：渲染 HTML 頁面
function doGet(e) {
  var template = HtmlService.createTemplateFromFile('Index');
  
  // 接收 URL 傳進來的 floor 參數 (例如 ?floor=1)
  template.currentFloor = (e && e.parameter && e.parameter.floor) ? e.parameter.floor : '';
  
  return template.evaluate()
    .setTitle('DEAR 午讀評分系統')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 2. 根據樓層獲取班別清單 (供前端選單使用)
function getClassesForFloor(floor) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("班別名單");
    if (!sheet) return [];
    
    var data = sheet.getDataRange().getValues();
    var classes = [];
    
    // 假設 A 欄為樓層，B 欄為班別名稱 (依實際情況調整標頭列)
    for (var i = 1; i < data.length; i++) {
      var rowFloor = String(data[i][0]).trim();
      var className = String(data[i][1]).trim();
      
      if (!floor || rowFloor === String(floor).trim()) {
        classes.push(className);
      }
    }
    return classes;
  } catch (error) {
    Logger.log("Error in getClassesForFloor: " + error.toString());
    return [];
  }
}

// 3. 處理前端提交的評分紀錄 (核心寫入邏輯)
function submitReadData(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("評分紀錄");
    
    if (!sheet) {
      return { success: false, message: "找不到「評分紀錄」工作表！" };
    }

    // A. 判斷評分歸屬日期 (優先使用補交日期)
    var recordDateStr;
    if (data.customDate && String(data.customDate).trim() !== "") {
      recordDateStr = String(data.customDate).trim(); // 使用補交日期 (YYYY-MM-DD)
    } else {
      var now = new Date();
      recordDateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd"); // 使用今日日期
    }

    // B. 自動計算「月份」欄位 (如 "8月")
    var dateParts = recordDateStr.split("-");
    var monthNum = parseInt(dateParts[1], 10);
    var monthLabel = monthNum + "月";

    // C. 真正按下提交的系統時間 (稽核用)
    var submitTimestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

    // D. 寫入 Google Sheet 試算表 (依序對齊你的 A~J 欄)
    sheet.appendRow([
      submitTimestamp,                           // A欄: 提交時間
      recordDateStr,                             // B欄: 紀錄日期 (YYYY-MM-DD)
      monthLabel,                                // C欄: 月份 (如 8月)
      data.floor || '',                          // D欄: 樓層
      data.className || '',                      // E欄: 班別
      data.topic || '',                          // F欄: 本月主題
      Number(data.baseScore) || 0,               // G欄: 基本分
      Number(data.extraScore) || 0,              // H欄: 額外分
      Number(data.totalScore) || 0,              // I欄: 總分
      data.isBackfill ? "【補交紀錄】" : "【即時紀錄】" // J欄: 備註
    ]);

    return { 
      success: true, 
      message: "提交成功！紀錄日期：" + recordDateStr + " (" + monthLabel + ")" 
    };

  } catch (error) {
    return { 
      success: false, 
      message: "寫入失敗：" + error.toString() 
    };
  }
}
