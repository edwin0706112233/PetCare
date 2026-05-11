const catNames = ["芬達", "77"];

function onEdit(e) {
  // 防呆：確保 e 存在
  if (!e) return; 

  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName();
  const row = range.getRow();

  // 1. 【守衛區】不符合條件的，通通提早結束
  const ignoreSheets = ["飼料清單", "飼料機發糧計劃", "Siri捷徑"];
  if (ignoreSheets.includes(sheetName)) return; 
  if (range.getColumn() !== 3) return;          
  if (sheet.getLastRow() !== row) return;       // 只在最後一列觸發

  // 2. 【共用動作】符合條件後，執行共用的設定日期
  setDate(sheet, row);

  // 3. 【任務分發】如果是吃飯紀錄，就把工作交給專門的 Function 去處理
  const targetCats = ["芬達吃飯", "77吃飯"];
  if (targetCats.includes(sheetName)) {
    handleCatFeeding(e.source, sheet, sheetName, row);
  }
}

// 新增 customDateStr 與 customTimeStr 作為可選參數
function setDate(sheet, row, customDateStr, customTimeStr) {
  let dateOnly, timeOnly;

  // 1. 判斷是否有前端傳來的日期 (YYYY-MM-DD)
  if (customDateStr) {
    const dParts = customDateStr.split('-');
    dateOnly = new Date(dParts[0], dParts[1] - 1, dParts[2]);
  } else {
    // 沒有的話，就抓伺服器當下日期 (給 onEdit 用)
    const today = new Date();
    dateOnly = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  }

  // 2. 判斷是否有前端傳來的時間 (HH:mm)
  if (customTimeStr) {
    const tParts = customTimeStr.split(':');
    timeOnly = new Date(0, 0, 0, tParts[0], tParts[1], 0);
  } else {
    // 沒有的話，就抓伺服器當下時間 (給 onEdit 用)
    const today = new Date();
    timeOnly = new Date(0, 0, 0, today.getHours(), today.getMinutes(), today.getSeconds());
  }
  
  // 3. 寫入資料 (使用陣列一次寫入 A、B 欄，效能最佳化)
  sheet.getRange(row, 1, 1, 2).setValues([[dateOnly, timeOnly]]);
  
  // 如果你的試算表 A、B 欄還沒手動設定格式，可以把下面這行解除註解
  // sheet.getRange(row, 1, 1, 2).setNumberFormats([["M/dd", "HH:mm"]]);
}

function setFoodWeight(sheet, row, weight) {
  sheet.getRange(row, 5).setValue(weight); 
}

function setWaterVolume(sheet, row, volume) {
  sheet.getRange(row, 7).setValue(volume); 
}

function doPost (e) {
  factory(e);
  
  return "succeeds";
}

function doGet(e) {
  // 1. 判斷是否有傳入 action 參數
  // 如果有 action，就走你原本的 API 路線 (回傳純文字)
  if (e && e.parameter && e.parameter.action) {
    const SpreadSheet = SpreadsheetApp.openById('15SooPjS1LNik2QPr6CYoIJfvaUV0GY6Fxfp6v8knKMo');
    var responseText = ""
    var action = e.parameter.action;
    
    switch (action) {
      case 'action1':
        responseText = getLastFeed(SpreadSheet);
        break;
      case 'action2':
        responseText = getTodayTotal(SpreadSheet);
        break;
      default:
        responseText = "找不到 action"
    }
    
    return ContentService.createTextOutput(responseText);
  }
  // 建立 HTML 模板
  var template = HtmlService.createTemplateFromFile('index');
  
  // ✨ 關鍵魔法：檢查網址列有沒有帶入 ?view=readonly 這個暗號
  template.isReadOnly = (e.parameter && e.parameter.view === 'readonly');
  
  // 輸出網頁
  return template.evaluate()
    .setTitle('🐾 喵星人健康管家')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
}

function resetCheckboxes() {
  // ===== 請在這裡修改工作表名稱和儲存格範圍 =====
  var sheetName = "餵食計畫";  // <--- 請改成您的工作表分頁名稱
  var rangeToResetFanta = "I9:I14";    // <--- 請改成您放置核取方塊的儲存格範圍
  var rangeToResetSeven = "N9:N14";    // <--- 請改成您放置核取方塊的儲存格範圍
  // ===========================================

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);

  if (sheet) {
    sheet.getRange(rangeToResetFanta).uncheck();
    sheet.getRange(rangeToResetSeven).uncheck();
  }
}

// function getFoodOptions() {
//   const ss = SpreadsheetApp.openById('15SooPjS1LNik2QPr6CYoIJfvaUV0GY6Fxfp6v8knKMo');
//   const sheet = ss.getSheetByName("飼料清單");
//   const data = sheet.getDataRange().getValues();
  
//   // 假設你的「濕食、乾飼料、副食罐」是寫在第一列 (Row 1)
//   const headers = data[0];
//   const wetCol = headers.indexOf("濕食");
//   const dryCol = headers.indexOf("乾飼料");
//   const subCol = headers.indexOf("副食罐");

//   let options = {
//     "濕食": [],
//     "乾飼料": [],
//     "副食罐": []
//   };

//   // 從第二列開始往下抓，把不是空白的名稱收進陣列
//   for (let i = 1; i < data.length; i++) {
//     if (wetCol !== -1 && data[i][wetCol]) options["濕食"].push(data[i][wetCol]);
//     if (dryCol !== -1 && data[i][dryCol]) options["乾飼料"].push(data[i][dryCol]);
//     if (subCol !== -1 && data[i][subCol]) options["副食罐"].push(data[i][subCol]);
//   }
  
//   return options;
// }
// 1. 取得初始化資料：一次把「貓咪名單」跟「飼料菜單」都抓給網頁
function getInitData() {
  const ss = SpreadsheetApp.openById('15SooPjS1LNik2QPr6CYoIJfvaUV0GY6Fxfp6v8knKMo');
  
  // -- 抓取貓咪名單 --
  const catSheet = ss.getSheetByName("貓咪名稱");
  if (!catSheet) throw new Error("找不到『貓咪名稱』分頁！");
  const catData = catSheet.getDataRange().getValues();
  let catNames = [];
  
  for (let i = 1; i < catData.length; i++) {
    if (catData[i][0]) { 
      catNames.push(catData[i][0].toString().trim());
    }
  }

  // -- 抓取飼料清單 --
  const foodSheet = ss.getSheetByName("飼料清單");
  const foodData = foodSheet.getDataRange().getValues();
  const headers = foodData[0];
  
  // ✨ 找出各個欄位的索引
  const wetCol = headers.indexOf("濕食");
  const dryCol = headers.indexOf("乾飼料");
  const subCol = headers.indexOf("副食罐");
  const liqCol = headers.indexOf("營養液"); // ✨ 增加這行

  // ✨ 初始化物件，加入「營養液」
  let foodOptions = { "濕食": [], "乾飼料": [], "副食罐": [], "營養液": [] };

  for (let i = 1; i < foodData.length; i++) {
    if (wetCol !== -1 && foodData[i][wetCol]) foodOptions["濕食"].push(foodData[i][wetCol]);
    if (dryCol !== -1 && foodData[i][dryCol]) foodOptions["乾飼料"].push(foodData[i][dryCol]);
    if (subCol !== -1 && foodData[i][subCol]) foodOptions["副食罐"].push(foodData[i][subCol]);
    if (liqCol !== -1 && foodData[i][liqCol]) foodOptions["營養液"].push(foodData[i][liqCol]); // ✨ 增加這行
  }
  
  return {
    cats: catNames,
    foods: foodOptions
  };
}

// 2. 寫入資料 (邏輯不變，函數名稱改為通用的 addCatRecord)
function addCatRecord(data) {
  const ss = SpreadsheetApp.openById('15SooPjS1LNik2QPr6CYoIJfvaUV0GY6Fxfp6v8knKMo'); 
  
  // ✨ 核心魔法：這裡直接用網頁傳來的貓咪名稱加上「吃飯」去找分頁
  const sheetName = data.catName + "吃飯";
  const sheet = ss.getSheetByName(sheetName); 
  
  if (!sheet) throw new Error("找不到『" + sheetName + "』分頁，請確認是否已建立！");

  for (let i = 0; i < data.records.length; i++) {
    let record = data.records[i];
    let rowData = [
      "", "", record.foodName, "", record.amount, "", record.water, record.notes, "", "", "", ""
    ];
    sheet.appendRow(rowData);
    
    // 呼叫你的 setDate 模組
    const lastRow = sheet.getLastRow();
    setDate(sheet, lastRow, data.date, data.time);
  }

  return "✅ 成功寫入 " + data.catName + " 的 " + data.records.length + " 筆紀錄！";
}

// 抓取散佈圖需要的時間數據 (支援濕食、乾飼料、副食罐)
// function getCatFeedingTimeStats(catName) {
//   const ss = SpreadsheetApp.openById('15SooPjS1LNik2QPr6CYoIJfvaUV0GY6Fxfp6v8knKMo');
//   const sheet = ss.getSheetByName(catName + "吃飯");
//   if (!sheet) return { error: "找不到表單" };

//   const data = sheet.getDataRange().getValues();
//   const validRows = data.slice(1).filter(row => row[0] !== "");

//   const now = new Date();
//   const cutoff = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 29);
//   cutoff.setHours(0, 0, 0, 0);

//   let wetData = [];
//   let dryData = [];
//   let subData = []; // ✨ 新增：用來裝「副食罐」的陣列

//   validRows.forEach(row => {
//     let dateObj = row[0];
//     let timeObj = row[1];
//     let type = String(row[3]).trim(); // D 欄：種類

//     let d = new Date(dateObj);
//     if (isNaN(d.getTime()) || d < cutoff) return; 

//     let dateStr = (d.getMonth() + 1) + "/" + d.getDate();

//     let hours = 0, mins = 0;
//     if (timeObj instanceof Date) {
//       hours = timeObj.getHours();
//       mins = timeObj.getMinutes();
//     } else if (typeof timeObj === 'string') {
//       let parts = timeObj.split(':');
//       if (parts.length >= 2) {
//         hours = parseInt(parts[0], 10);
//         mins = parseInt(parts[1], 10);
//       }
//     }
//     let timeValue = hours + (mins / 60);
//     let dataPoint = { x: dateStr, y: timeValue };

//     // ✨ 根據種類分裝到對應的陣列中
//     if (type === "濕食") wetData.push(dataPoint);
//     else if (type === "乾飼料") dryData.push(dataPoint);
//     else if (type === "副食罐") subData.push(dataPoint); 
//   });

//   let labels = [];
//   for (let i = 29; i >= 0; i--) {
//     let tempD = new Date();
//     tempD.setDate(tempD.getDate() - i);
//     labels.push((tempD.getMonth() + 1) + "/" + tempD.getDate());
//   }

//   // ✨ 回傳時，把 sub 也一起打包送給前端
//   return { labels: labels, wet: wetData, dry: dryData, sub: subData };
// }

// 取得貓咪過去 30 天的體重數據 (同日多筆取平均)
// function getCatWeightStats(catName) {
//   const ss = SpreadsheetApp.openById('15SooPjS1LNik2QPr6CYoIJfvaUV0GY6Fxfp6v8knKMo');
//   const sheet = ss.getSheetByName("體重");
//   if (!sheet) return { error: "找不到體重表單" };

//   const data = sheet.getDataRange().getValues();
  
//   const validRows = data.slice(2).filter(row => {
//     // 強制將 row[3] 轉為字串並去除空白，再與 catName 比較
//     return row[0] !== "" && String(row[3]).trim() === String(catName).trim();
//   });

//   const now = new Date();
//   const cutoff = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 29);
//   cutoff.setHours(0, 0, 0, 0);

//   let dailyWeights = {};

//   validRows.forEach(row => {
//     let dateObj = row[0];
//     let weight = Number(row[2]); 
//     if (isNaN(weight)) return;

//     let d = new Date(dateObj);
//     if (isNaN(d.getTime()) || d < cutoff) return; 

//     let dateStr = (d.getMonth() + 1) + "/" + d.getDate();

//     if (!dailyWeights[dateStr]) {
//       dailyWeights[dateStr] = { sum: 0, count: 0 };
//     }
    
//     dailyWeights[dateStr].sum += weight;
//     dailyWeights[dateStr].count += 1;
//   });

//   let labels = [];
//   let weightData = [];

//   for (let i = 29; i >= 0; i--) {
//     let tempD = new Date();
//     tempD.setDate(tempD.getDate() - i);
//     let dateStr = (tempD.getMonth() + 1) + "/" + tempD.getDate();
//     labels.push(dateStr);

//     if (dailyWeights[dateStr]) {
//       let avgWeight = dailyWeights[dateStr].sum / dailyWeights[dateStr].count;
//       weightData.push(Number(avgWeight.toFixed(2)));
//     } else {
//       weightData.push(null);
//     }
//   }

//   return { labels: labels, weight: weightData };
// }

// 取得貓咪過去 30 天的上廁所(便便)次數
// function getCatPoopStats(catName) {
//   const ss = SpreadsheetApp.openById('15SooPjS1LNik2QPr6CYoIJfvaUV0GY6Fxfp6v8knKMo');
//   const sheet = ss.getSheetByName("貓砂");
//   if (!sheet) return { error: "找不到貓砂表單" };

//   const data = sheet.getDataRange().getValues();
  
//   // ✨ 關鍵：你的資料從第 4 行開始，所以陣列索引是 3 (slice(3))
//   const validRows = data.slice(3).filter(row => {
//     // 確保有日期，且 C 欄 (index 2) 的動物名稱符合
//     return row[0] !== "" && String(row[2]).trim() === String(catName).trim();
//   });

//   const now = new Date();
//   const cutoff = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 29);
//   cutoff.setHours(0, 0, 0, 0);

//   // 用 Object 來暫存每天的次數
//   let dailyCounts = {};

//   validRows.forEach(row => {
//     let dateObj = row[0];
//     let d = new Date(dateObj);
//     if (isNaN(d.getTime()) || d < cutoff) return; // 略過 30 天前

//     // 格式化為 MM/DD
//     let dateStr = (d.getMonth() + 1) + "/" + d.getDate();

//     // 次數累加
//     if (!dailyCounts[dateStr]) {
//       dailyCounts[dateStr] = 0;
//     }
//     dailyCounts[dateStr]++;
//   });

//   let labels = [];
//   let counts = [];

//   // 產生過去 30 天的完整 X 軸
//   for (let i = 29; i >= 0; i--) {
//     let tempD = new Date();
//     tempD.setDate(tempD.getDate() - i);
//     let dateStr = (tempD.getMonth() + 1) + "/" + tempD.getDate();
    
//     labels.push(dateStr);
//     // 如果這天有便便紀錄就塞入次數，沒有就是 0
//     counts.push(dailyCounts[dateStr] || 0);
//   }

//   return { labels: labels, counts: counts };
// }

// 取得兩貓共用的水碗消耗量
// function getSharedWaterStats() {
//   const ss = SpreadsheetApp.openById('15SooPjS1LNik2QPr6CYoIJfvaUV0GY6Fxfp6v8knKMo');
//   const sheet = ss.getSheetByName("喝水");
//   if (!sheet) return { error: "找不到喝水表單" };

//   const data = sheet.getDataRange().getValues();
  
//   // 你的資料從第 4 行開始，索引為 3
//   const validRows = data.slice(3).filter(row => row[0] !== "");

//   // 用來儲存每天的總消耗水量
//   let dailyWater = {};
  
//   // 暫存最近一次「開始」的重量
//   let lastStartWeight = null;

//   // 由上往下依序處理歷史紀錄
//   validRows.forEach(row => {
//     let dateObj = row[0];
//     let weight = Number(row[2]); // C 欄：重量(g)
//     let type = String(row[3]).trim(); // D 欄：種類 (開始/結束)

//     let d = new Date(dateObj);
//     if (isNaN(d.getTime())) return;

//     if (type === "開始") {
//       // 記錄開始重量
//       lastStartWeight = weight;
//     } 
//     else if (type === "結束") {
//       // 確保有先記錄過「開始」，且開始重量大於結束重量
//       if (lastStartWeight !== null && lastStartWeight >= weight) {
//         let consumed = lastStartWeight - weight; // 算出消耗量
        
//         let dateStr = (d.getMonth() + 1) + "/" + d.getDate();
//         if (!dailyWater[dateStr]) dailyWater[dateStr] = 0;
        
//         // 將消耗量累加到該日期中 (如果一天內換了兩次水，就會加總)
//         dailyWater[dateStr] += consumed;
        
//         // 算完後將開始重量清空，等待下一次的「開始」
//         lastStartWeight = null; 
//       }
//     }
//   });

//   let labels = [];
//   let waterData = [];

//   // 產生過去 30 天的完整 X 軸
//   for (let i = 29; i >= 0; i--) {
//     let tempD = new Date();
//     tempD.setDate(tempD.getDate() - i);
//     let dateStr = (tempD.getMonth() + 1) + "/" + tempD.getDate();
    
//     labels.push(dateStr);
//     waterData.push(dailyWater[dateStr] || 0); // 沒有消耗就補 0
//   }

//   return { labels: labels, data: waterData };
// }

// ================= 圖表資料 API (支援自訂日期區間) =================

// 1. 取得熱量與水量
function getCatMonthlyIntakeStats(catName, startDateStr, endDateStr) {
  const ss = SpreadsheetApp.openById('15SooPjS1LNik2QPr6CYoIJfvaUV0GY6Fxfp6v8knKMo');
  const sheet = ss.getSheetByName(catName + "吃飯");
  if (!sheet) return { error: "找不到表單" };
  
  const data = sheet.getDataRange().getValues().slice(1);
  const start = new Date(startDateStr); start.setHours(0,0,0,0);
  const end = new Date(endDateStr); end.setHours(23,59,59,999);

  let dailyData = {};
  data.forEach(row => {
    let d = new Date(row[0]);
    if (isNaN(d.getTime()) || d < start || d > end) return;
    let dateStr = (d.getMonth() + 1) + "/" + d.getDate();

    if (!dailyData[dateStr]) dailyData[dateStr] = { cal: 0, water: 0 };
    dailyData[dateStr].water += (Number(row[6]) || 0) + (Number(row[9]) || 0); // 加水 + 罐頭水
    
    // ✨ 關鍵修改：從 row[10] (K欄) 改為 row[8] (I欄: 熱量)
    dailyData[dateStr].cal += (Number(row[8]) || 0); 
  });

  let labels = [], calories = [], water = [];
  for (let tempD = new Date(start); tempD <= end; tempD.setDate(tempD.getDate() + 1)) {
    let dateStr = (tempD.getMonth() + 1) + "/" + tempD.getDate();
    labels.push(dateStr);
    calories.push(dailyData[dateStr] ? dailyData[dateStr].cal : 0);
    water.push(dailyData[dateStr] ? dailyData[dateStr].water : 0);
  }
  return { labels, calories, water };
}

// 2. 取得進食時間散佈圖
function getCatFeedingTimeStats(catName, startDateStr, endDateStr) {
  const ss = SpreadsheetApp.openById('15SooPjS1LNik2QPr6CYoIJfvaUV0GY6Fxfp6v8knKMo');
  const sheet = ss.getSheetByName(catName + "吃飯");
  if (!sheet) return { error: "找不到表單" };

  const data = sheet.getDataRange().getValues().slice(1);
  const start = new Date(startDateStr); start.setHours(0,0,0,0);
  const end = new Date(endDateStr); end.setHours(23,59,59,999);

  let wetData = [], dryData = [], subData = [], liqData = []; // ✨ 新增 liqData
  data.forEach(row => {
    let timeObj = row[1];
    let type = String(row[3]).trim(); 
    if (!timeObj || timeObj === "") return; 

    let d = new Date(row[0]);
    if (isNaN(d.getTime()) || d < start || d > end) return; 
    let dateStr = (d.getMonth() + 1) + "/" + d.getDate();

    let hours = 0, mins = 0;
    if (timeObj instanceof Date) { hours = timeObj.getHours(); mins = timeObj.getMinutes(); } 
    else if (typeof timeObj === 'string') { let parts = timeObj.split(':'); if (parts.length >= 2) { hours = parseInt(parts[0], 10); mins = parseInt(parts[1], 10); } }
    
    let timeValue = hours + (mins / 60);
    let dataPoint = { x: dateStr, y: timeValue };

    if (type === "濕食") wetData.push(dataPoint);
    else if (type === "乾飼料") dryData.push(dataPoint);
    else if (type === "副食罐") subData.push(dataPoint); 
    else if (type === "營養液") liqData.push(dataPoint); // ✨ 新增判斷
  });

  let labels = [];
  for (let tempD = new Date(start); tempD <= end; tempD.setDate(tempD.getDate() + 1)) {
    labels.push((tempD.getMonth() + 1) + "/" + tempD.getDate());
  }
  return { labels, wet: wetData, dry: dryData, sub: subData, liq: liqData }; // ✨ 傳回 liq
}

// 3. 取得體重變化
function getCatWeightStats(catName, startDateStr, endDateStr) {
  const ss = SpreadsheetApp.openById('15SooPjS1LNik2QPr6CYoIJfvaUV0GY6Fxfp6v8knKMo');
  const sheet = ss.getSheetByName("體重");
  if (!sheet) return { error: "找不到體重表單" };

  const data = sheet.getDataRange().getValues().slice(2).filter(row => row[0] !== "" && String(row[3]).trim() === String(catName).trim());
  const start = new Date(startDateStr); start.setHours(0,0,0,0);
  const end = new Date(endDateStr); end.setHours(23,59,59,999);

  let dailyWeights = {};
  data.forEach(row => {
    let weight = Number(row[2]); 
    if (isNaN(weight)) return;
    let d = new Date(row[0]);
    if (isNaN(d.getTime()) || d < start || d > end) return; 

    let dateStr = (d.getMonth() + 1) + "/" + d.getDate();
    if (!dailyWeights[dateStr]) dailyWeights[dateStr] = { sum: 0, count: 0 };
    dailyWeights[dateStr].sum += weight;
    dailyWeights[dateStr].count += 1;
  });

  let labels = [], weightData = [];
  for (let tempD = new Date(start); tempD <= end; tempD.setDate(tempD.getDate() + 1)) {
    let dateStr = (tempD.getMonth() + 1) + "/" + tempD.getDate();
    labels.push(dateStr);
    if (dailyWeights[dateStr]) {
      weightData.push(Number((dailyWeights[dateStr].sum / dailyWeights[dateStr].count).toFixed(2)));
    } else {
      weightData.push(null);
    }
  }
  return { labels, weight: weightData };
}

// 4. 取得便便次數
// ================= ✨ 貓砂圖表三分類升級版 (正常/拉肚/亂尿) =================
function getCatPoopStats(catName, startStr, endStr) {
  try {
    const ss = SpreadsheetApp.openById('15SooPjS1LNik2QPr6CYoIJfvaUV0GY6Fxfp6v8knKMo');
    // 雖然 sheet 名字叫貓砂，但我們現在把它當作「健康異常事件表」來用！
    const sheet = ss.getSheetByName("貓砂");
    if (!sheet) return { labels: [], normal: [], diarrhea: [], litter: [], vomit: [] };

    const data = sheet.getDataRange().getValues();
    const startDate = new Date(startStr);
    const endDate = new Date(endStr);
    endDate.setHours(23, 59, 59, 999);

    let stats = {};

    for (let i = 1; i < data.length; i++) {
      let row = data[i];
      let rowDate = new Date(row[0]);
      if (isNaN(rowDate.getTime())) continue;

      let rowCat = String(row[2] || "").trim();
      let rowNotes = String(row[3] || "").trim();

      if (rowCat === catName && rowDate >= startDate && rowDate <= endDate) {
        let dateKey = Utilities.formatDate(rowDate, "Asia/Taipei", "M/d");
        // ✨ 加入 vomit (嘔吐) 初始值
        if (!stats[dateKey]) stats[dateKey] = { normal: 0, diarrhea: 0, litter: 0, vomit: 0 };
        
        // ✨ 分類邏輯：加入嘔吐判斷
        if (rowNotes.includes("嘔吐") || rowNotes.includes("吐")) {
          stats[dateKey].vomit += 1;
        } else if (rowNotes.includes("亂尿尿")) {
          stats[dateKey].litter += 1;
        } else if (rowNotes.includes("拉肚子") || rowNotes.includes("軟便") || rowNotes.includes("水便")) {
          stats[dateKey].diarrhea += 1;
        } else if (rowNotes.includes("正常") || rowNotes.includes("大便")) {
          stats[dateKey].normal += 1;
        } else if (!rowNotes.includes("尿尿")) {
          stats[dateKey].normal += 1; 
        }
      }
    }

    let labels = [];
    let normalCounts = [];
    let diarrheaCounts = [];
    let litterCounts = [];
    let vomitCounts = []; // ✨ 新增嘔吐陣列
    
    let curr = new Date(startDate);
    while (curr <= endDate) {
      let dateKey = Utilities.formatDate(curr, "Asia/Taipei", "M/d");
      labels.push(dateKey);
      if (stats[dateKey]) {
        normalCounts.push(stats[dateKey].normal);
        diarrheaCounts.push(stats[dateKey].diarrhea);
        litterCounts.push(stats[dateKey].litter);
        vomitCounts.push(stats[dateKey].vomit); // ✨ 推入嘔吐數據
      } else {
        normalCounts.push(0);
        diarrheaCounts.push(0);
        litterCounts.push(0);
        vomitCounts.push(0);
      }
      curr.setDate(curr.getDate() + 1);
    }

    // ✨ 回傳 5 個東西，包含 vomit
    return { labels: labels, normal: normalCounts, diarrhea: diarrheaCounts, litter: litterCounts, vomit: vomitCounts };

  } catch (error) {
    return { labels: [], normal: [], diarrhea: [], litter: [], vomit: [] };
  }
}

// 5. 取得共用水碗
function getSharedWaterStats(startDateStr, endDateStr) {
  const ss = SpreadsheetApp.openById('15SooPjS1LNik2QPr6CYoIJfvaUV0GY6Fxfp6v8knKMo');
  const sheet = ss.getSheetByName("喝水");
  if (!sheet) return { error: "找不到喝水表單" };

  const data = sheet.getDataRange().getValues().slice(3).filter(row => row[0] !== "");
  const start = new Date(startDateStr); start.setHours(0,0,0,0);
  const end = new Date(endDateStr); end.setHours(23,59,59,999);

  let dailyWater = {};
  let lastStartWeight = null;

  data.forEach(row => {
    let weight = Number(row[2]); 
    let type = String(row[3]).trim(); 
    let d = new Date(row[0]);
    if (isNaN(d.getTime())) return;

    if (type === "開始") {
      lastStartWeight = weight;
    } else if (type === "結束") {
      if (lastStartWeight !== null && lastStartWeight >= weight) {
        // 如果結算當下落在我們選的日期區間內，才算進去
        if (d >= start && d <= end) {
          let consumed = lastStartWeight - weight; 
          let dateStr = (d.getMonth() + 1) + "/" + d.getDate();
          dailyWater[dateStr] = (dailyWater[dateStr] || 0) + consumed;
        }
        lastStartWeight = null; 
      }
    }
  });

  let labels = [], waterData = [];
  for (let tempD = new Date(start); tempD <= end; tempD.setDate(tempD.getDate() + 1)) {
    let dateStr = (tempD.getMonth() + 1) + "/" + tempD.getDate();
    labels.push(dateStr);
    waterData.push(dailyWater[dateStr] || 0); 
  }
  return { labels, data: waterData };
}

// 取得指定貓咪「今日」的已攝取水量與熱量
function getTodayIntakeSummary(catName) {
  const ss = SpreadsheetApp.openById('15SooPjS1LNik2QPr6CYoIJfvaUV0GY6Fxfp6v8knKMo');
  const sheet = ss.getSheetByName(catName + "吃飯");
  if (!sheet) return { water: 0, calories: 0 };

  const data = sheet.getDataRange().getValues().slice(1); 
  const today = new Date();
  const todayStr = today.getFullYear() + '/' + (today.getMonth() + 1) + '/' + today.getDate();

  let totalWater = 0;
  let totalCalories = 0;

  data.forEach(row => {
    let d = new Date(row[0]);
    if (isNaN(d.getTime())) return;
    
    let rowDateStr = d.getFullYear() + '/' + (d.getMonth() + 1) + '/' + d.getDate();
    
    if (rowDateStr === todayStr) {
      let addWater = Number(row[6]) || 0;        // G欄: 加水(cc)
      let cannedWater = Number(row[9]) || 0;     // J欄: 罐頭含水量 (cc)
      
      // ✨ 關鍵修改：從 row[10] (K欄) 改為 row[8] (I欄: 熱量)
      let foodCalories = Number(row[8]) || 0;    

      totalWater += (addWater + cannedWater);
      totalCalories += foodCalories;
    }
  });

  return { 
    water: Math.round(totalWater),     
    calories: Math.round(totalCalories) 
  };
}

// 取得指定貓咪最近的常用紀錄 (排除完全重複的組合，各分類取前 2 筆)
function getRecentRecords(catName) {
  const ss = SpreadsheetApp.openById('15SooPjS1LNik2QPr6CYoIJfvaUV0GY6Fxfp6v8knKMo');
  const sheet = ss.getSheetByName(catName + "吃飯");
  if (!sheet) return { error: "找不到表單" };

  const data = sheet.getDataRange().getValues();
  const validRows = data.slice(1);

  // 準備用來裝捷徑的物件
  let recent = { "濕食": [], "乾飼料": [], "副食罐": [], "營養液": [] };
  let seen = { "濕食": new Set(), "乾飼料": new Set(), "副食罐": new Set(), "營養液": new Set() };

  // ✨ 從最後一行 (最新紀錄) 往上找
  for (let i = validRows.length - 1; i >= 0; i--) {
    let row = validRows[i];
    let foodName = String(row[2]).trim(); // C欄: 食物名
    let type = String(row[3]).trim();     // D欄: 種類
    let amount = Number(row[4]) || 0;     // E欄: 數量
    let water = Number(row[6]) || 0;      // G欄: 加水
    let notes = String(row[7]).trim();    // H欄: 備註

    // 如果這個種類還沒收集滿 2 筆，且食物名不是空的
    if (recent[type] && recent[type].length < 2 && foodName !== "") {
      // 用字串把這個組合包起來當作唯一識別碼 (避免連續兩天吃一模一樣的佔用兩個按鈕)
      let key = `${foodName}_${amount}_${water}_${notes}`;
      
      if (!seen[type].has(key)) {
        seen[type].add(key);
        recent[type].push({
          foodName: foodName,
          amount: amount,
          water: water,
          notes: notes
        });
      }
    }

    // 提早結束機制：如果三個分類都各找到 2 筆了，就停止掃描 (節省效能)
    if (recent["濕食"].length >= 2 && recent["乾飼料"].length >= 2 && recent["副食罐"].length >= 2 && recent["營養液"].length >= 2) {
      break;
    }
  }

  return recent;
}

// ================= ✨ 刪除歷史紀錄系統 =================

// 1. 取得最近 10 筆紀錄 (供 UI 顯示與選擇)
function getRecentRecordsForDeletion(catName) {
  const ss = SpreadsheetApp.openById('15SooPjS1LNik2QPr6CYoIJfvaUV0GY6Fxfp6v8knKMo');
  const sheet = ss.getSheetByName(catName + "吃飯");
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  let records = [];
  
  for (let i = data.length - 1; i >= 1; i--) {
    let row = data[i];
    let d = new Date(row[0]);
    if (isNaN(d.getTime())) continue; 

    // ✨ 修復：讓刪除清單裡的日期顯示標準的 yyyy/MM/dd，不會再只有 3/10
    let dateStr = Utilities.formatDate(d, "Asia/Taipei", "yyyy/MM/dd");
    
    let timeObj = row[1];
    let timeStr = "";
    if (timeObj instanceof Date) {
      timeStr = Utilities.formatDate(timeObj, "Asia/Taipei", "HH:mm");
    } else {
      timeStr = String(timeObj);
    }

    records.push({
      rowIndex: i + 1, 
      date: dateStr,
      time: timeStr,
      foodName: String(row[2]).trim() || String(row[3]).trim(), 
      amount: Number(row[4]) || 0,
      water: Number(row[6]) || 0,
      notes: String(row[7]).trim()
    });

    if (records.length >= 10) break;
  }
  return records;
}

// 2. 刪除指定的行號，並自動將下方資料往上移
function deleteRecord(catName, rowIndex) {
  const ss = SpreadsheetApp.openById('15SooPjS1LNik2QPr6CYoIJfvaUV0GY6Fxfp6v8knKMo');
  const sheet = ss.getSheetByName(catName + "吃飯");
  if (!sheet) return { success: false, message: "找不到表單" };
  
  try {
    // ✨ 這行是魔法：刪除整列，後面的資料會自動往前遞補！
    sheet.deleteRow(rowIndex);
    return { success: true, message: "紀錄已成功刪除！" };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ================= ✨ 重大事件 (里程碑) 紀錄功能 =================

// 1. 儲存重大事件到獨立表單
function saveMilestone(catName, dateStr, text) {
  try {
    const ss = SpreadsheetApp.openById('15SooPjS1LNik2QPr6CYoIJfvaUV0GY6Fxfp6v8knKMo');
    let sheet = ss.getSheetByName("重大事件紀錄");
    
    // 如果找不到這個分頁，就自動建立一個
    if (!sheet) {
      sheet = ss.insertSheet("重大事件紀錄");
      sheet.appendRow(["日期", "貓咪名稱", "事件說明"]);
      // 凍結第一列
      sheet.setFrozenRows(1);
    }
    
    // 寫入資料
    sheet.appendRow([new Date(dateStr), catName, text]);
    return { success: true };
  } catch (e) {
    throw new Error(e.toString());
  }
}

// 2. 獲取圖表區間內的重大事件，傳給前端畫紅線
function getMilestones(catName, startDateStr, endDateStr) {
  try {
    const ss = SpreadsheetApp.openById('15SooPjS1LNik2QPr6CYoIJfvaUV0GY6Fxfp6v8knKMo');
    const sheet = ss.getSheetByName("重大事件紀錄");
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    let start = new Date(startDateStr); start.setHours(0,0,0,0);
    let end = new Date(endDateStr); end.setHours(23,59,59,999);
    
    let milestones = [];
    
    for (let i = 1; i < data.length; i++) {
      let rowDate = new Date(data[i][0]);
      let rowCat = String(data[i][1]).trim();
      let text = String(data[i][2]).trim();
      
      // 篩選：日期在區間內，且符合目前的貓咪
      if (rowCat === catName && rowDate >= start && rowDate <= end) {
        let formattedDate = Utilities.formatDate(rowDate, "Asia/Taipei", "M/d");
        milestones.push({
          date: formattedDate,
          text: text
        });
      }
    }
    
    return milestones;
  } catch (e) {
    return { error: e.toString() };
  }
}

// ================= ✨ 修改歷史紀錄 =================
function updateRecordData(catName, rowIndex, amount, water, notes) {
  try {
    const ss = SpreadsheetApp.openById('15SooPjS1LNik2QPr6CYoIJfvaUV0GY6Fxfp6v8knKMo');
    const sheet = ss.getSheetByName(catName + "吃飯");
    if (!sheet) return { success: false, message: "找不到該貓咪的表單" };

    // ⚠️ 欄位對應（請確認這與你的 Google Sheet 一致）：
    // 假設 第6欄(F)是重量，第7欄(G)是水量，第8欄(H)是備註
    sheet.getRange(rowIndex, 6).setValue(amount); 
    sheet.getRange(rowIndex, 7).setValue(water);  
    sheet.getRange(rowIndex, 8).setValue(notes);  

    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}