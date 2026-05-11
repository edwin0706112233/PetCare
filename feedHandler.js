function handleCatFeeding(ss, sheet, catName, row) {
  // 取得剛剛編輯的那一列，第 4 欄的食物種類 (假設這是字串，例如 "濕食" 或 "乾飼料")
  const feedType = sheet.getRange(row, 4).getValue(); 
  const feedSheet = getSheet(CONFIG.SHEETS.FEED_LIST);

  // 【效能優化核心】不要一次一次取，一次把 F2:H5 (列2~5，欄6~8) 範圍的資料讀進陣列
  // configData[0] 就是 Row 2 (芬達濕食)，以此類推
  const configData = feedSheet.getRange(2, 6, 10, 3).getValues();

  // 【查表法】建立一個貓咪專屬的「設定字典」，對應你原本的 if/else
  // 結構：{ "貓名": { "食物名稱": { weight: 重量, water: 水量 } } }
  const feedConfig = {
    [CONFIG.SHEETS.FENDA_FEED]: {
      [configData[0][0]]: { weight: configData[0][1], water: configData[0][2] }, // 芬達濕食 (Row 2)
      [configData[1][0]]: { weight: configData[1][1], water: configData[1][2] },  // 芬達乾糧 (Row 3)
      [configData[2][0]]: { weight: configData[2][1], water: configData[2][2] }  // 芬達營養液 (Row 4)
    },
    [CONFIG.SHEETS.SEVENSEVEN_FEED]: {
      [configData[3][0]]: { weight: configData[3][1], water: configData[3][2] }, // 77濕食 (Row 5)
      [configData[4][0]]: { weight: configData[4][1], water: configData[4][2] }  // 77乾糧 (Row 6)
    }
  };

  // 根據現在是哪隻貓，去拿牠的專屬菜單
  const catMenu = feedConfig[catName];
  if (!catMenu) return; // 找不到該貓咪菜單就結束

  // 根據輸入的食物名稱，找出對應的重量和水量
  const foodSetting = catMenu[feedType];

  if (foodSetting) {
    // 找到對應食物，寫入重量與水量
    setFoodWeight(sheet, row, foodSetting.weight);
    setWaterVolume(sheet, row, foodSetting.water);
  } else if (catName === CONFIG.SHEETS.SEVENSEVEN_FEED) {
    // 找不到對應食物，且是你原本設定 77 的防呆邏輯
    setFoodWeight(sheet, row, -1);
    setWaterVolume(sheet, row, -1);
  }
}

// ================= ✨ 食慾評鑑系統 (P欄專用版) =================

// 1. 取得最近 3 筆內「已紀錄食物，且時間在昨天(含)以內，但還未評分」的清單
// ================= ✨ 食慾評鑑系統 (P欄專用版 + 智慧過期過濾) =================
function getUnratedRecords(catName) {
  try {
    const ss = getSpreadsheet();
    const sheet = getSheet(catName + "吃飯");
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    let unrated = [];
    
    // ✨ 計算「昨天」的凌晨 00:00 的時間基準點
    const yesterday = getYesterdayStart();

    const startIndex = Math.max(1, data.length - 30);

    for (let i = startIndex; i < data.length; i++) {
      let row = data[i];
      let dateVal = row[0];
      let foodName = String(row[2] || "").trim();
      let appetite = String(row[15] || "").trim(); // P 欄

      // 確認 dateVal 是一個有效的日期物件
      let recordDate = new Date(dateVal);
      let isValidDate = (recordDate instanceof Date && !isNaN(recordDate));

      // ✨ 判斷：有填食物、沒填食慾、是有效日期、而且「日期必須大於或等於昨天」！
      if (foodName !== "" && appetite === "" && isValidDate && recordDate >= yesterday) {
        let dateStr = formatSimpleDate(recordDate);
        let timeObj = row[1];
        let timeStr = formatTime(timeObj);

        unrated.push({
          rowIndex: i + 1,
          date: dateStr,
          time: timeStr,
          foodName: foodName,
          amount: row[4] || 0
        });
      }
    }
    // 反轉陣列，讓最新的一餐排在最上面
    return unrated.reverse(); 
  } catch (e) {
    return [];
  }
}
// 2. 使用者點擊表情符號後，更新該筆紀錄的食慾分數
function updateAppetiteScore(catName, rowIndex, score) {
  try {
    const ss = getSpreadsheet();
    const sheet = getSheet(catName + "吃飯");
    if (!sheet) return { success: false, message: "找不到表單" };

    // ✨ 改這裡！直接在該行的第 16 欄 (P 欄) 寫入分數
    sheet.getRange(rowIndex, 16).setValue(score);
    return { success: true };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ================= ✨ 取得食慾統計圖表資料 =================

// ================= ✨ 取得食慾統計圖表資料 (已修正時間軸對齊) =================

function getCatAppetiteStats(catName, startDateStr, endDateStr) {
  try {
    const ss = getSpreadsheet();
    const sheet = getSheet(catName + "吃飯");
    if (!sheet) return { error: "找不到表單" };

    const data = sheet.getDataRange().getValues();
    let stats = {};

    let start = new Date(startDateStr); start.setHours(0,0,0,0);
    let end = new Date(endDateStr); end.setHours(23,59,59,999);

    for (let i = 1; i < data.length; i++) {
      let row = data[i];
      let dateVal = row[0];
      if (!(dateVal instanceof Date)) continue;

      if (dateVal >= start && dateVal <= end) {
        // ✨ 改用與其他圖表一致的 M/d 格式
        let dateStr = formatSimpleDate(dateVal);
        let appetite = row[15]; 
        
        if (appetite !== "" && !isNaN(appetite)) {
          let score = parseFloat(appetite);
          if (!stats[dateStr]) stats[dateStr] = { totalScore: 0, count: 0 };
          stats[dateStr].totalScore += score;
          stats[dateStr].count++;
        }
      }
    }

    let labels = [];
    let avgScores = [];

    // ✨ 強制補齊區間內的所有日期，確保有空檔的日子紅線也能畫得出來！
    for (let tempD = new Date(start); tempD <= end; tempD.setDate(tempD.getDate() + 1)) {
      let dStr = formatSimpleDate(tempD);
      labels.push(dStr);
      if (stats[dStr]) {
        avgScores.push((stats[dStr].totalScore / stats[dStr].count).toFixed(1)); 
      } else {
        avgScores.push(null); // 當天沒紀錄就留空，圖表會自動跨過去連線
      }
    }

    return { labels: labels, scores: avgScores };
  } catch (e) {
    return { error: e.toString() };
  }
}