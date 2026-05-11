function factory(e) {
  // 防呆：確保有收到內容
  if (!e || !e.postData || !e.postData.contents) return;
  const params = JSON.parse(e.postData.contents);

  // 1. 處理吃飯紀錄 (來自手機)
  if (params.source == "mobile") {
    // 簡化判斷式
    const sheetName = getCatSheetName(params.user);
    eatFood(params, sheetName);
    return; // 執行完提早結束
  } 
  
  // 2. 處理其他事件 (使用 switch 讓路由更清晰)
  switch (params.event) {
    case "貓砂":
      noteLitterBoxWeight(params);
      break;
    case "體重":
      noteCatWeight(params);
      break;
    case "喝水":
      noteCatDrinkWater(params);
      break;
  }
}

function eatFood(params, sheetName) {
  const SpreadSheet = getSpreadsheet();
  const Sheet = SpreadSheet.getSheetByName(sheetName); 
  if (!Sheet) return;

  const configSheet = SpreadSheet.getSheetByName("Siri捷徑");
  const configData = configSheet.getDataRange().getValues();
  const rowsToWrite = [];

  const catConfigStartColumn = (sheetName == "77吃飯") ? 5 : 0;
  const targetKey = params.event;

  // 使用 Date 物件，之後寫入時讓 Sheets 自動套用已設定好的格式
  const now = getNowInTimezone();
  const dateOnly = getTodayStart();
  const timeOnly = createTime(now.getHours(), now.getMinutes(), now.getSeconds());

  // 從 i = 1 開始 (跳過標題)
  for (let i = 1; i < configData.length; i++) {
    let row = configData[i];
    let configKey = row[0 + catConfigStartColumn];
    
    if (configKey == targetKey) {
      let itemName = row[1 + catConfigStartColumn];
      let weight = (row[2 + catConfigStartColumn] || 0);
      let water = (row[3 + catConfigStartColumn] || 0);

      // 順序對應: A:日期, B:時間, C:食物名, D:種類(留空), E:數量, F:剩餘(留空), G:加水
      rowsToWrite.push([dateOnly, timeOnly, itemName, "", weight, "", water]);
    }
  }

  // 批次寫入資料
  if (rowsToWrite.length > 0) {
    const LastRow = Sheet.getLastRow();
    Sheet.getRange(LastRow + 1, 1, rowsToWrite.length, rowsToWrite[0].length).setValues(rowsToWrite);
    // 註：只要 A、B 欄在試算表上已經設定過日期/時間格式，這裡就不需要再頻繁呼叫 setNumberFormat
  }
}

// ----------------------------------------------------
// 以下三個 note 開頭的 Function，全部用 appendRow 濃縮成 1 次寫入！
// 寫入陣列 [日期, 時間, 參數1, 參數2...]
// ----------------------------------------------------

function noteLitterBoxWeight(params) {
  const SpreadSheet = getSpreadsheet();
  const Sheet = SpreadSheet.getSheetByName(CONFIG.SHEETS.LITTER);
  
  // ✨ 順手幫你把日期時間升級成最安全的台北時區格式，避免 Siri 紀錄時跨天跑版
  const now = getNowInTimezone();
  const dateStr = formatDate(now);
  const timeStr = formatTime(now);
  
  // 取得 Siri 傳來的貓咪名字
  const catName = toString(params.name);
  
  // ✨ 核心邏輯：接收 Siri 傳來的「狀態」。如果 Siri 沒傳，就預設為「正常大便」
  // 這樣你原本只喊「紀錄芬達貓砂」的舊捷徑也不會壞掉！
  let litterType = toString(params.type) || "正常大便"; 
  
  // 依序寫入: A欄(日期), B欄(時間), C欄(動物), D欄(備註=狀態)
  Sheet.appendRow([dateStr, timeStr, catName, litterType]); 
}

// ================= ✨ 貓砂與歷史紀錄系統 (日期格式終極修復版) =================

// 1. 寫入貓砂紀錄
function addLitterRecord(catName, date, time, type, notes) {
  const ss = getSpreadsheet();
  const sheet = getSheet(CONFIG.SHEETS.LITTER);
  if (!sheet) throw new Error("找不到『貓砂』試算表");
  
  // ✨ 修復核心：把前端的 "2026-03-23" 強制轉成試算表最愛的 "2026/03/23"
  let formattedDate = toString(date).replace(/-/g, '/');
  
  let finalNotes = toString(type);
  if (notes) finalNotes += " - " + toString(notes);
  
  sheet.appendRow([formattedDate, time, catName, finalNotes]);
  return "貓砂紀錄已儲存！";
}

function noteCatWeight(params) {
  const SpreadSheet = getSpreadsheet();
  const Sheet = SpreadSheet.getSheetByName(CONFIG.SHEETS.WEIGHT);

  const now = getNowInTimezone();
  const dateOnly = getTodayStart();
  const timeOnly = createTime(now.getHours(), now.getMinutes(), now.getSeconds());

  // 依序寫入: A欄(日期), B欄(時間), C欄(體重), D欄(使用者)
  Sheet.appendRow([dateOnly, timeOnly, params.weight, params.user]); 
}

function noteCatDrinkWater(params) {
  const SpreadSheet = getSpreadsheet();
  const Sheet = SpreadSheet.getSheetByName(CONFIG.SHEETS.WATER);

  const now = getNowInTimezone();
  const dateOnly = getTodayStart();
  const timeOnly = createTime(now.getHours(), now.getMinutes(), now.getSeconds());

  // 依序寫入: A欄(日期), B欄(時間), C欄(水量), D欄(種類)
  Sheet.appendRow([dateOnly, timeOnly, params.weight, params.type]); 
}

// ----------------------------------------------------
// 以下為查詢資料的 Function (批次讀取優化)
// ----------------------------------------------------

function getLastFeed(SpreadSheet) {
  const sheets = [
    SpreadSheet.getSheetByName(CONFIG.SHEETS.FENDA_FEED),
    SpreadSheet.getSheetByName(CONFIG.SHEETS.SEVENSEVEN_FEED)
  ];
  const sheetNames = [CONFIG.SHEETS.FENDA_FEED, CONFIG.SHEETS.SEVENSEVEN_FEED];
  const catNames = sheetNames.map(name => getCatDisplayName(name));
  let responses = [];

  for (let i = 0; i < sheets.length; i++) {
    let sheet = sheets[i];
    if (!sheet) continue;

    let lastRow = sheet.getLastRow(); 
    if (lastRow < 2) continue; // 確保有資料

    // 【效能優化】：一次把 B 欄到 E 欄 (2~5) 的資料讀出來
    let rowData = sheet.getRange(lastRow, 2, 1, 4).getDisplayValues()[0];
    
    let timeValue = toString(rowData[0]); // B欄: 時間
    let foodName = toString(rowData[1]);  // C欄: 食物名
    let weight = toString(rowData[3]);    // E欄: 重量 (索引3)

    responses.push(`${catNames[i]}最後進食時間是 ${timeValue}，吃的是 ${foodName}，重量 ${weight} 克。`);
  }

  return "報告主人：" + responses.join(" ");
}
function getTodayTotal(SpreadSheet) {
  const sheets = [
    SpreadSheet.getSheetByName(CONFIG.SHEETS.FENDA_FEED),
    SpreadSheet.getSheetByName(CONFIG.SHEETS.SEVENSEVEN_FEED)
  ];
  const sheetNames = [CONFIG.SHEETS.FENDA_FEED, CONFIG.SHEETS.SEVENSEVEN_FEED];
  const catNames = sheetNames.map(name => getCatDisplayName(name));
  let responses = [];

  // 取得今天 00:00:00 的時間物件進行比較
  const todayStart = getTodayStart().getTime();

  for (let i = 0; i < sheets.length; i++) {
    let sheet = sheets[i];
    if (!sheet) continue;

    let lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      responses.push(`${catNames[i]} 今天尚無進食紀錄。`);
      continue;
    }

    // 依據您的新欄位順序：
    // A:日期, B:時間, C:食物名, D:種類, E:數量, F:剩餘, G:加水(索引6), H:備註, I:熱量(索引8), J:罐頭含水量(索引9)
    let data = sheet.getRange(2, 1, lastRow - 1, 10).getValues(); 
    let dailyCalories = 0;
    let dailyWater = 0;

    for (let j = 0; j < data.length; j++) {
      let row = data[j];
      let rowDate = row[0];

      if (rowDate instanceof Date) {
        // 將該列日期轉為毫秒數，確保它是今天 00:00 之後的資料
        let rowDateTime = new Date(rowDate.getFullYear(), rowDate.getMonth(), rowDate.getDate()).getTime();

        if (rowDateTime === todayStart) {
          let calories = toNumber(row[8], 0);     // I 欄：熱量
          let waterAdded = toNumber(row[6], 0);   // G 欄：加水
          let waterInCan = toNumber(row[9], 0);   // J 欄：罐頭含水量
          
          dailyCalories += calories;
          dailyWater += (waterAdded + waterInCan);
        }
      }
    }

    responses.push(`${catNames[i]}：今天攝取了${dailyCalories.toFixed(1)} 大卡，喝了 ${dailyWater.toFixed(1)} cc的水`);
  }

  return "報告主人：" + responses.join("\n");
}