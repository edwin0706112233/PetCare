function handleCatFeeding(ss, sheet, catName, row) {
  const feedType = sheet.getRange(row, 4).getValue(); 
  const feedSheet = getSheet(CONFIG.SHEETS.FEED_LIST);

  const configData = feedSheet.getRange(2, 6, 10, 3).getValues();
  const feedConfig = {
    [CONFIG.SHEETS.FENDA_FEED]: {
      [configData[0][0]]: { weight: configData[0][1], water: configData[0][2] },
      [configData[1][0]]: { weight: configData[1][1], water: configData[1][2] },
      [configData[2][0]]: { weight: configData[2][1], water: configData[2][2] }
    },
    [CONFIG.SHEETS.SEVENSEVEN_FEED]: {
      [configData[3][0]]: { weight: configData[3][1], water: configData[3][2] },
      [configData[4][0]]: { weight: configData[4][1], water: configData[4][2] }
    }
  };

  const catMenu = feedConfig[catName];
  if (!catMenu) return;

  const foodSetting = catMenu[feedType];

  if (foodSetting) {
    setFoodWeight(sheet, row, foodSetting.weight);
    setWaterVolume(sheet, row, foodSetting.water);
  } else if (catName === CONFIG.SHEETS.SEVENSEVEN_FEED) {
    // 找不到對應食物，且是你原本設定 77 的防呆邏輯
    setFoodWeight(sheet, row, -1);
    setWaterVolume(sheet, row, -1);
  }
}


function getUnratedRecords(catName) {
  try {
    const ss = getSpreadsheet();
    const sheet = getSheet(catName + "吃飯");
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    let unrated = [];
    
    const yesterday = getYesterdayStart();
    const startIndex = Math.max(1, data.length - 30);

    for (let i = startIndex; i < data.length; i++) {
      let row = data[i];
      let dateVal = row[0];
      let foodName = String(row[2] || "").trim();
      let appetite = String(row[15] || "").trim();

      let recordDate = new Date(dateVal);
      let isValidDate = (recordDate instanceof Date && !isNaN(recordDate));

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
    return unrated.reverse();
  } catch (e) {
    return [];
  }
}

function updateAppetiteScore(catName, rowIndex, score) {
  try {
    const ss = getSpreadsheet();
    const sheet = getSheet(catName + "吃飯");
    if (!sheet) return { success: false, message: "找不到表單" };

    sheet.getRange(rowIndex, 16).setValue(score);
    return { success: true };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}


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

    for (let tempD = new Date(start); tempD <= end; tempD.setDate(tempD.getDate() + 1)) {
      let dStr = formatSimpleDate(tempD);
      labels.push(dStr);
      if (stats[dStr]) {
        avgScores.push((stats[dStr].totalScore / stats[dStr].count).toFixed(1)); 
      } else {
        avgScores.push(null);
      }
    }

    return { labels: labels, scores: avgScores };
  } catch (e) {
    return { error: e.toString() };
  }
}