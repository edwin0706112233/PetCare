function factory(e) {
  if (!e || !e.postData || !e.postData.contents) return;
  const params = JSON.parse(e.postData.contents);

  if (params.source == "mobile") {
    const sheetName = getCatSheetName(params.user);
    eatFood(params, sheetName);
    return;
  } 
  
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

  const now = getNowInTimezone();
  const dateOnly = getTodayStart();
  const timeOnly = createTime(now.getHours(), now.getMinutes(), now.getSeconds());

  for (let i = 1; i < configData.length; i++) {
    let row = configData[i];
    let configKey = row[0 + catConfigStartColumn];
    
    if (configKey == targetKey) {
      let itemName = row[1 + catConfigStartColumn];
      let weight = (row[2 + catConfigStartColumn] || 0);
      let water = (row[3 + catConfigStartColumn] || 0);

      rowsToWrite.push([dateOnly, timeOnly, itemName, "", weight, "", water]);
    }
  }

  if (rowsToWrite.length > 0) {
    const LastRow = Sheet.getLastRow();
    Sheet.getRange(LastRow + 1, 1, rowsToWrite.length, rowsToWrite[0].length).setValues(rowsToWrite);
  }
}


function noteLitterBoxWeight(params) {
  const SpreadSheet = getSpreadsheet();
  const Sheet = SpreadSheet.getSheetByName(CONFIG.SHEETS.LITTER);
  
  const now = getNowInTimezone();
  const dateStr = formatDate(now);
  const timeStr = formatTime(now);
  
  const catName = toString(params.name);
  let litterType = toString(params.type) || "正常大便";
  
  Sheet.appendRow([dateStr, timeStr, catName, litterType]); 
}

// ================= ✨ 貓砂與歷史紀錄系統 (日期格式終極修復版) =================

// 1. 寫入貓砂紀錄
function addLitterRecord(catName, date, time, type, notes) {
  const ss = getSpreadsheet();
  const sheet = getSheet(CONFIG.SHEETS.LITTER);
  if (!sheet) throw new Error("找不到『貓砂』試算表");
  

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
    if (lastRow < 2) continue;

    let rowData = sheet.getRange(lastRow, 2, 1, 4).getDisplayValues()[0];
    
    let timeValue = toString(rowData[0]);
    let foodName = toString(rowData[1]);
    let weight = toString(rowData[3]);

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