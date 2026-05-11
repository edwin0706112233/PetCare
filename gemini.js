// ================= ✨ AI 獸醫助理系統 =================

function checkAvailableModels() {
  const API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  const url = "https://generativelanguage.googleapis.com/v1beta/models?key=" + API_KEY;
  
  const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  Logger.log(response.getContentText());
}


// 1. 幫 AI 整理病歷資料 (將吃飯、貓砂與重大事件紀錄合併並按時間排序)
function getRawDataForAI(catName, startStr, endStr) {
  const ss = SpreadsheetApp.openById('15SooPjS1LNik2QPr6CYoIJfvaUV0GY6Fxfp6v8knKMo'); // 你的試算表 ID
  const startDate = new Date(startStr);
  const endDate = new Date(endStr);
  endDate.setHours(23, 59, 59, 999);

  let logData = [];

  // 抓取吃飯資料
  const foodSheet = ss.getSheetByName(catName + "吃飯");
  if (foodSheet) {
    let fData = foodSheet.getDataRange().getValues();
    for (let i = 1; i < fData.length; i++) {
      let rowDate = new Date(fData[i][0]);
      if (isNaN(rowDate.getTime())) continue;
      if (rowDate >= startDate && rowDate <= endDate) {
        let dateStr = Utilities.formatDate(rowDate, "Asia/Taipei", "yyyy/MM/dd");
        let timeObj = fData[i][1];
        let timeStr = (timeObj instanceof Date) ? Utilities.formatDate(timeObj, "Asia/Taipei", "HH:mm") : String(timeObj);
        let foodName = String(fData[i][2] || fData[i][3]).trim();
        let amount = fData[i][4] || 0;
        let water = fData[i][6] || 0;
        let notes = String(fData[i][7] || "").trim();
        logData.push({ 
          timestamp: rowDate.getTime() + parseInt(timeStr.replace(':','')), // 用來排序
          text: `[${dateStr} ${timeStr}] 飲食: 吃了 ${foodName} ${amount}g, 喝水 ${water}ml。備註: ${notes}` 
        });
      }
    }
  }

  // 抓取貓砂與異常資料
  const litterSheet = ss.getSheetByName("貓砂");
  if (litterSheet) {
    let lData = litterSheet.getDataRange().getValues();
    for (let i = 1; i < lData.length; i++) {
      let rowDate = new Date(lData[i][0]);
      if (isNaN(rowDate.getTime())) continue;
      let rowCat = String(lData[i][2] || "").trim();
      if (rowCat === catName && rowDate >= startDate && rowDate <= endDate) {
        let dateStr = Utilities.formatDate(rowDate, "Asia/Taipei", "yyyy/MM/dd");
        let timeObj = lData[i][1];
        let timeStr = (timeObj instanceof Date) ? Utilities.formatDate(timeObj, "Asia/Taipei", "HH:mm") : String(timeObj);
        let notes = String(lData[i][3] || "").trim();
        logData.push({ 
          timestamp: rowDate.getTime() + parseInt(timeStr.replace(':','')), // 用來排序
          text: `[${dateStr} ${timeStr}] 排泄/異常: ${notes}` 
        });
      }
    }
  }

  // ✨ 新增：抓取重大事件資料
  const milestoneSheet = ss.getSheetByName("重大事件紀錄");
  if (milestoneSheet) {
    let mData = milestoneSheet.getDataRange().getValues();
    // 根據你提供的格式：A欄(0)=日期, B欄(1)=貓咪名稱, C欄(2)=事件說明
    for (let i = 1; i < mData.length; i++) {
      let rowDate = new Date(mData[i][0]);
      if (isNaN(rowDate.getTime())) continue;
      let rowCat = String(mData[i][1] || "").trim();
      
      if (rowCat === catName && rowDate >= startDate && rowDate <= endDate) {
        let dateStr = Utilities.formatDate(rowDate, "Asia/Taipei", "yyyy/MM/dd");
        // 如果重大事件有包含時間，我們把它抓出來，沒有的話預設會是 00:00 或 08:00
        let timeStr = Utilities.formatDate(rowDate, "Asia/Taipei", "HH:mm");
        let event = String(mData[i][2] || "").trim();
        
        logData.push({ 
          timestamp: rowDate.getTime() + parseInt(timeStr.replace(':','')), // 用同樣的邏輯排序
          text: `[${dateStr} ${timeStr}] ⚠️ 重大醫療/生活事件: ${event}` 
        });
      }
    }
  }

  // 將資料依照時間先後順序排列，讓 AI 能夠看出時間軸變化
  logData.sort((a, b) => a.timestamp - b.timestamp);

  let finalLog = logData.map(item => item.text).join("\n");
  return finalLog || "這段期間內沒有找到任何紀錄。";
}

// 2. 呼叫 Gemini API 進行分析
function askGeminiAboutCat(catName, question, startStr, endStr) {
  // 🔥🔥🔥 請把下面這行的引號內容，換成你的 API Key 🔥🔥🔥
  const API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  
  if (API_KEY.includes("請在這裡貼上")) {
    throw new Error("請先到 .gs 檔案中填寫你的 Gemini API Key！");
  }

  // 去抓出這個區間的完整貓咪病歷 (現在已經包含重大事件了)
  const logText = getRawDataForAI(catName, startStr, endStr);

  // ✨ 組合給 AI 的終極提示詞 (Prompt) -> 加上了重大事件的提醒
  const prompt = `你現在是一位專業、細心且充滿同理心的貓咪獸醫助理。
請根據以下這段期間內「${catName}」的真實健康紀錄，來回答主人的問題。
如果紀錄中有出現「亂尿尿」、「拉肚子」、「嘔吐」或是食水量急遽減少，請特別提醒主人注意。
⚠️ 另外，如果紀錄中有出現「⚠️ 重大醫療/生活事件」（例如手術、用藥等），請務必將這些事件與貓咪前後的食慾、排泄狀況進行交叉比對與綜合分析。
請用溫柔、口語化且排版清晰的繁體中文回答。如果沒有資料，請直接告訴主人這段期間沒有紀錄。

【${catName} 的健康紀錄】
${logText}

【主人的問題】
${question}`;

  const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key=" + API_KEY;
  const payload = {
    "contents": [{ "parts": [{ "text": prompt }] }],
    "generationConfig": { "temperature": 0.3 } // 讓 AI 的回答更偏向冷靜客觀的醫療分析
  };

  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  const response = UrlFetchApp.fetch(url, options);
  const json = JSON.parse(response.getContentText());

  if (json.error) {
    throw new Error("AI 呼叫失敗: " + json.error.message);
  }

  return json.candidates[0].content.parts[0].text;
}

// // 1. 幫 AI 整理病歷資料 (將吃飯與貓砂紀錄合併並按時間排序)
// function getRawDataForAI(catName, startStr, endStr) {
//   const ss = SpreadsheetApp.openById('15SooPjS1LNik2QPr6CYoIJfvaUV0GY6Fxfp6v8knKMo'); // 你的試算表 ID
//   const startDate = new Date(startStr);
//   const endDate = new Date(endStr);
//   endDate.setHours(23, 59, 59, 999);

//   let logData = [];

//   // 抓取吃飯資料
//   const foodSheet = ss.getSheetByName(catName + "吃飯");
//   if (foodSheet) {
//     let fData = foodSheet.getDataRange().getValues();
//     for (let i = 1; i < fData.length; i++) {
//       let rowDate = new Date(fData[i][0]);
//       if (isNaN(rowDate.getTime())) continue;
//       if (rowDate >= startDate && rowDate <= endDate) {
//         let dateStr = Utilities.formatDate(rowDate, "Asia/Taipei", "yyyy/MM/dd");
//         let timeObj = fData[i][1];
//         let timeStr = (timeObj instanceof Date) ? Utilities.formatDate(timeObj, "Asia/Taipei", "HH:mm") : String(timeObj);
//         let foodName = String(fData[i][2] || fData[i][3]).trim();
//         let amount = fData[i][4] || 0;
//         let water = fData[i][6] || 0;
//         let notes = String(fData[i][7] || "").trim();
//         logData.push({ 
//           timestamp: rowDate.getTime() + parseInt(timeStr.replace(':','')), // 用來排序
//           text: `[${dateStr} ${timeStr}] 飲食: 吃了 ${foodName} ${amount}g, 喝水 ${water}ml。備註: ${notes}` 
//         });
//       }
//     }
//   }

//   // 抓取貓砂與異常資料
//   const litterSheet = ss.getSheetByName("貓砂");
//   if (litterSheet) {
//     let lData = litterSheet.getDataRange().getValues();
//     for (let i = 1; i < lData.length; i++) {
//       let rowDate = new Date(lData[i][0]);
//       if (isNaN(rowDate.getTime())) continue;
//       let rowCat = String(lData[i][2] || "").trim();
//       if (rowCat === catName && rowDate >= startDate && rowDate <= endDate) {
//         let dateStr = Utilities.formatDate(rowDate, "Asia/Taipei", "yyyy/MM/dd");
//         let timeObj = lData[i][1];
//         let timeStr = (timeObj instanceof Date) ? Utilities.formatDate(timeObj, "Asia/Taipei", "HH:mm") : String(timeObj);
//         let notes = String(lData[i][3] || "").trim();
//         logData.push({ 
//           timestamp: rowDate.getTime() + parseInt(timeStr.replace(':','')), // 用來排序
//           text: `[${dateStr} ${timeStr}] 排泄/異常: ${notes}` 
//         });
//       }
//     }
//   }

//   // 將資料依照時間先後順序排列，讓 AI 能夠看出時間軸變化
//   logData.sort((a, b) => a.timestamp - b.timestamp);

//   let finalLog = logData.map(item => item.text).join("\n");
//   return finalLog || "這段期間內沒有找到任何紀錄。";
// }

// // 2. 呼叫 Gemini API 進行分析
// function askGeminiAboutCat(catName, question, startStr, endStr) {
//   // 🔥🔥🔥 請把下面這行的引號內容，換成你的 API Key 🔥🔥🔥
//   const API_KEY = "AIzaSyB7lyB1tEIuT5umDglcQA1WUo8OmMsGsPQ";
  
//   if (API_KEY.includes("請在這裡貼上")) {
//     throw new Error("請先到 .gs 檔案中填寫你的 Gemini API Key！");
//   }

//   // 去抓出這個區間的完整貓咪病歷
//   const logText = getRawDataForAI(catName, startStr, endStr);

//   // 組合給 AI 的終極提示詞 (Prompt)
//   const prompt = `你現在是一位專業、細心且充滿同理心的貓咪獸醫助理。
// 請根據以下這段期間內「${catName}」的真實健康紀錄，來回答主人的問題。
// 如果紀錄中有出現「亂尿尿」、「拉肚子」、「嘔吐」或是食水量急遽減少，請特別提醒主人注意。
// 請用溫柔、口語化且排版清晰的繁體中文回答。如果沒有資料，請直接告訴主人這段期間沒有紀錄。

// 【${catName} 的健康紀錄】
// ${logText}

// 【主人的問題】
// ${question}`;

//   const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key=" + API_KEY;
//   const payload = {
//     "contents": [{ "parts": [{ "text": prompt }] }],
//     "generationConfig": { "temperature": 0.3 } // 讓 AI 的回答更偏向冷靜客觀的醫療分析
//   };

//   const options = {
//     "method": "post",
//     "contentType": "application/json",
//     "payload": JSON.stringify(payload),
//     "muteHttpExceptions": true
//   };

//   const response = UrlFetchApp.fetch(url, options);
//   const json = JSON.parse(response.getContentText());

//   if (json.error) {
//     throw new Error("AI 呼叫失敗: " + json.error.message);
//   }

//   return json.candidates[0].content.parts[0].text;
// }