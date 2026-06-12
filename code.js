/**
 * Sherry Aerial 薪資核對系統 - 2026 最終除錯穩定版
 * 更新重點：
 * 1. 精準對齊所有欄位，防止前端網頁崩潰。
 * 2. 加入「保留確認狀態」機制：重新執行計算時，不會洗掉老師已確認的狀態與時間。
 * 3. 強化字串與空白防呆處理，確保明細 100% 抓取。
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('💰 薪資系統功能')
      .addItem('執行當月薪資計算', 'calculateSalaries')
      .addToUi();
}

// --- 1. 主計算邏輯 ---
function calculateSalaries() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summarySheet = ss.getSheetByName("每月更新總表");
  const ladderSheet = ss.getSheetByName("薪項表"); 
  const resultSheet = ss.getSheetByName("計算結果");
  const specialSheet = ss.getSheetByName("特別課");
  
  if (!summarySheet || !ladderSheet || !resultSheet) {
    SpreadsheetApp.getUi().alert("錯誤：找不到標籤頁（每月更新總表、薪項表 或 計算結果）");
    return;
  }

  // 抓取總表資料 (A到K欄，共11欄，確保涵蓋到 I欄的盈利)
  const lastRow = summarySheet.getLastRow();
  if (lastRow < 2) return;
  const summaryRange = summarySheet.getRange(2, 1, lastRow - 1, 11);
  const summaryValues = summaryRange.getValues();
  const ladderValues = ladderSheet.getDataRange().getValues();
  
  // 建立特別課索引
  let specialMap = {};
  if (specialSheet) {
    const specData = specialSheet.getDataRange().getValues();
    for (let i = 1; i < specData.length; i++) {
      let d = specData[i][0];
      let dStr = (d instanceof Date) ? Utilities.formatDate(d, "GMT+8", "yyyy-MM-dd") : String(d || "").trim();
      let key = dStr + "_" + String(specData[i][1] || "").trim() + "_" + String(specData[i][2] || "").trim();
      specialMap[key] = true;
    }
  }

  let teacherSummary = {}; 

  for (let i = 0; i < summaryValues.length; i++) {
    let row = summaryValues[i];
    let courseName = String(row[1] || "").trim(); 
    let dateObj = row[2];
    let attStr = String(row[3] || "").trim(); 
    let instructorStr = String(row[5] || "").trim(); // F欄: 指導者
    let income = Number(String(row[7] || "").replace(/,/g, '')) || 0; // H欄: 課程收入

    // 處理場地租借
    if (courseName.includes("場地租借")) {
      summaryValues[i][6] = 0; // G欄: 鐘點費
      summaryValues[i][8] = 0; // I欄: 盈利
      continue;
    }
    
    // 排除空指導者與共用帳號
    if (!instructorStr || instructorStr === "共用帳號") {
      summaryValues[i][6] = 0; 
      summaryValues[i][8] = income; 
      continue;
    }
    
    let actualCount = attStr.includes('/') ? parseInt(attStr.split('/')[0], 10) : parseInt(attStr, 10);
    if (isNaN(actualCount)) actualCount = 0;

    // 分割多位老師
    let teachers = instructorStr.split(/[&,、/]+/).map(t => t.trim()).filter(t => t !== "");
    if (teachers.length === 0) continue;
    
    let dateStr = (dateObj instanceof Date) ? Utilities.formatDate(dateObj, "GMT+8", "yyyy-MM-dd") : String(dateObj || "");
    let isSpecial = specialMap[dateStr + "_" + courseName + "_" + instructorStr];

    let totalClassFee = 0;
    if (isSpecial) {
      totalClassFee = income * 0.6;
    } else {
      totalClassFee = getPreciseFee(teachers[0], courseName, actualCount, ladderValues);
    }

    let profit = income - totalClassFee;

    // 拆分老師鐘點費與盈利
    teachers.forEach(tName => {
      if (!teacherSummary[tName]) teacherSummary[tName] = { base: 0, fixed: 0, profit: 0 };
      teacherSummary[tName].base += (totalClassFee / teachers.length);
      teacherSummary[tName].profit += (profit / teachers.length);
    });

    summaryValues[i][6] = Math.round(totalClassFee); 
    summaryValues[i][8] = Math.round(profit); 
  }
  
  // 寫回總表
  summaryRange.setValues(summaryValues);
  
  // 補足固定項目
  ladderValues.forEach(r => {
    let name = String(r[0] || "").trim();
    let type = String(r[2] || "").trim();
    if (name !== "" && name !== "預設值") {
      if (["固定月薪", "固定加給", "固定扣項"].includes(type)) {
        if (!teacherSummary[name]) teacherSummary[name] = { base: 0, fixed: 0, profit: 0 };
        let val = Number(String(r[4] || "").replace(/,/g, '')) || 0;
        if (type === "固定扣項") val = -Math.abs(val);
        teacherSummary[name].fixed += val;
      }
    }
  });

  // 更新結果表並取得最終薪資字典
  let finalTotals = updateResultSheet(teacherSummary, resultSheet);
  
  // 同步更新給雪莉的格式
  updateSherryFormat(finalTotals);

  SpreadsheetApp.getUi().alert("✅ 薪資計算與同步完成！\n\n已成功結算盈利、更新總表，並同步至「給雪莉的格式」。");
}

// --- 2. 核心比對函數 ---
function getPreciseFee(tName, course, count, ladder) {
  let myRules = ladder.filter(r => tName.includes(String(r[0] || "").trim()) && String(r[0] || "").trim() !== "");
  let activeRules = myRules.filter(r => ["標準時薪", "人數階梯"].includes(String(r[2] || "").trim()));
  let fee = null;

  if (activeRules.length > 0) {
    let matched = activeRules.filter(r => String(r[1] || "").trim() !== "" && String(r[1] || "").trim() !== "(留空)" && course.includes(String(r[1] || "").trim()));
    if (matched.length === 0) matched = activeRules.filter(r => String(r[1] || "").trim() === "" || String(r[1] || "").trim() === "(留空)");

    if (matched.length > 0) {
      let type = String(matched[0][2]).trim();
      if (type === "標準時薪") {
        fee = Number(matched[0][4]) || 0;
      } else if (type === "人數階梯") {
        matched.sort((a, b) => Number(b[3]) - Number(a[3]));
        for (let r of matched) {
          if (count >= Number(r[3])) { fee = Number(r[4]); break; }
        }
        if (fee === null) fee = Number(matched[matched.length - 1][4]); 
      }
    }
  }

  if (fee === null) {
    let defaultRules = ladder.filter(r => String(r[0] || "").trim() === "預設值" && ["標準時薪", "人數階梯"].includes(String(r[2] || "").trim()));
    let defMatched = defaultRules.filter(r => String(r[1] || "").trim() === "" || String(r[1] || "").trim() === "(留空)");
    if (defMatched.length > 0) {
      let type = String(defMatched[0][2]).trim();
      if (type === "標準時薪") fee = Number(defMatched[0][4]) || 0;
      else if (type === "人數階梯") {
        defMatched.sort((a, b) => Number(b[3]) - Number(a[3]));
        for (let r of defMatched) { if (count >= Number(r[3])) { fee = Number(r[4]); break; } }
        if (fee === null) fee = Number(defMatched[defMatched.length - 1][4]);
      }
    }
  }
  return fee !== null ? fee : 700; // 找不到對應時預設 700
}

// --- 3. 更新計算結果表 (加入保留狀態防呆) ---
function updateResultSheet(teacherSummary, resultSheet) {
  // 先備份原本的確認狀態與時間
  let existingStates = {};
  const lastResRow = resultSheet.getLastRow();
  if (lastResRow > 1) {
    const oldData = resultSheet.getRange(2, 1, lastResRow - 1, 9).getValues();
    oldData.forEach(r => {
      let n = String(r[0] || "").trim();
      if (n) {
        existingStates[n] = {
          status: String(r[7] || "").trim(), // H欄
          time: r[8] || ""                   // I欄
        };
      }
    });
  }

  resultSheet.clear();
  resultSheet.appendRow(["老師姓名", "鐘點費小計", "獎金比例", "獎金金額", "固定津貼/扣項", "應領總薪資", "盈利", "確認狀態", "確認時間"]);
  
  let finalTotals = {}; 

  for (let name in teacherSummary) {
    let base = Math.round(teacherSummary[name].base);
    let bonusRate = base >= 30000 ? 0.05 : base >= 20000 ? 0.04 : base >= 15000 ? 0.03 : 0;
    let bonus = Math.round(base * bonusRate);
    let fixed = Math.round(teacherSummary[name].fixed);
    let total = base + bonus + fixed;
    let profit = Math.round(teacherSummary[name].profit);
    
    finalTotals[name] = total; 

    // 恢復該老師原本的確認狀態，若沒有則預設「待確認」
    let currentStatus = (existingStates[name] && existingStates[name].status) ? existingStates[name].status : "待確認";
    let currentTime = (existingStates[name] && existingStates[name].time) ? existingStates[name].time : "";

    resultSheet.appendRow([name, base, (bonusRate*100)+"%", bonus, fixed, total, profit, currentStatus, currentTime]);
  }
  return finalTotals;
}

// --- 4. 更新「給雪莉的格式」分頁 ---
function updateSherryFormat(finalTotals) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sherrySheet = ss.getSheetByName("給雪莉的格式");
  if (!sherrySheet) return; 

  const teacherList = getTeacherList(); 
  const range = sherrySheet.getRange(1, 1, sherrySheet.getLastRow() || 1, 2);
  const values = range.getValues();

  for (let i = 0; i < values.length; i++) {
    let cellName = String(values[i][0] || "").trim();
    if (!cellName) continue;

    if (finalTotals.hasOwnProperty(cellName)) {
      values[i][1] = finalTotals[cellName];
    } else if (teacherList.includes(cellName)) {
      values[i][1] = 0; // 若為老師但當月無薪資則填 0
    }
  }
  range.setValues(values);
}

// --- 5. Web App 接口 (加強防崩潰) ---
function doGet(e) {
  const action = e.parameter.action;
  try {
    let result;
    if (action === "getTeachers") { result = getTeacherList(); } 
    else if (action === "login") { result = loginAndFetch(e.parameter.name, e.parameter.password); } 
    else if (action === "confirm") { result = submitConfirmation(e.parameter.name); }
    
    return ContentService.createTextOutput(JSON.stringify({success: true, data: result}))
                         .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({success: false, message: error.message}))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}

function getTeacherList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pwdSheet = ss.getSheetByName("密碼表");
  if (!pwdSheet) return [];
  const data = pwdSheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1).map(r => String(r[0] || "").trim()).filter(n => n !== "");
}

function loginAndFetch(name, password) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pwdSheet = ss.getSheetByName("密碼表");
  const resultSheet = ss.getSheetByName("計算結果");
  const summarySheet = ss.getSheetByName("每月更新總表");

  // 1. 驗證身分
  const auth = pwdSheet.getDataRange().getValues().find(r => String(r[0] || "").trim() === name && String(r[1] || "").trim() === password);
  if (!auth) throw new Error("姓名或密碼錯誤");

  // 2. 抓取結果表資料 (H欄對應 index 7)
  let resData = resultSheet.getDataRange().getValues().find(r => String(r[0] || "").trim() === name);
  if (!resData) {
    resData = [name, 0, "0%", 0, 0, 0, 0, "待確認", ""]; 
  }

  const cleanName = name.trim();
  const summaryData = summarySheet.getDataRange().getValues();

  // 3. 抓取明細
  const classes = summaryData.filter((r, idx) => {
    if (idx === 0) return false; 
    let instructorStr = String(r[5] || "").trim(); // F欄
    if (!instructorStr || instructorStr === "共用帳號") return false;

    // 寬鬆比對機制：避免 Emoji 或是空格造成的漏抓
    return instructorStr.includes(cleanName) || cleanName.includes(instructorStr);
  }).map(r => {
    let instructorStr = String(r[5] || "").trim();
    let teachers = instructorStr.split(/[&,、/]+/).map(t => t.trim()).filter(t => t !== "");
    
    let displayFee = Number(r[6]) || 0; // G欄：指導者鐘點費
    
    // 如果是合開課程，前台明細自動除以人數顯示
    if (teachers.length > 1 && displayFee > 0) {
      displayFee = Math.round(displayFee / teachers.length);
    }

    let dateStr = "";
    if (r[2] instanceof Date) {
      dateStr = Utilities.formatDate(r[2], "GMT+8", "MM/dd HH:mm");
    } else if (r[2]) {
      dateStr = String(r[2]).replace("2026-", "").substring(0, 11);
    }

    return {
      date: dateStr,
      title: String(r[1] || "").trim(), // B欄
      status: String(r[3] || "").trim(), // D欄
      fee: displayFee    
    };
  });

  return { 
    name: name, 
    base: resData[1], 
    bonus: resData[3], 
    fixed: resData[4], 
    total: resData[5], 
    classes: classes, 
    isConfirmed: String(resData[7] || "").trim() === "已確認" 
  };
}

function submitConfirmation(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("計算結果");
  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0] || "").trim() === name.trim()) {
      sheet.getRange(i + 1, 8).setValue("已確認");   // 寫入 H 欄
      sheet.getRange(i + 1, 9).setValue(new Date()); // 寫入 I 欄
      return true;
    }
  }
  return false;
}
