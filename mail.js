function generateAndEmailPDF(catName, startDate, endDate, emails, chartsBase64, customSubject, includeAI) {
  try {
    const emailSubject = customSubject ? customSubject : `🐾 【${catName}】健康分析報告 (${startDate} ~ ${endDate})`;
    
    const pdfFormalTitle = `🐾 ${catName} 的專屬健康分析報告`;
    
    const emailList = emails.split(',').map(e => e.trim()).filter(e => e !== "").join(',');

    let aiSectionHtml = "";
    if (includeAI) {
      const API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') || CONFIG.GEMINI_API_KEY;
      
      const logText = getRawDataForAI(catName, startDate, endDate);
      const prompt = `你是一位專業且細心的貓咪獸醫助理。這是一份 ${catName} 從 ${startDate} 到 ${endDate} 的健康數據：\n${logText}\n請針對這些數據寫一段綜合分析摘要（250字以內），指出值得注意的異常（如嘔吐、亂尿尿、水量變少），並給出專業建議。請用繁體中文，語氣溫和客觀。`;
      
      const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key=" + API_KEY;
      const payload = { "contents": [{ "parts": [{ "text": prompt }] }], "generationConfig": { "temperature": 0.2 } };
      const options = { "method": "post", "contentType": "application/json", "payload": JSON.stringify(payload), "muteHttpExceptions": true };
      
      const response = UrlFetchApp.fetch(url, options);
      const json = JSON.parse(response.getContentText());
      
      let aiText = "AI 無法取得分析資料。";
      if (!json.error && json.candidates && json.candidates[0]) {
         aiText = json.candidates[0].content.parts[0].text;
      }
      
      const formattedAiText = aiText.replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>').replace(/\n/g, '<br>');

      aiSectionHtml = `
        <div style="background: #f3e5f5; border-left: 5px solid #673ab7; padding: 20px; border-radius: 8px; margin-bottom: 35px; font-size: 15px; color: #333; line-height: 1.6;">
          <h3 style="margin-top: 0; color: #512da8; border-bottom: 1px dashed #d1c4e9; padding-bottom: 8px; margin-bottom: 12px;">🤖 獸醫助理綜合分析</h3>
          ${formattedAiText}
        </div>
      `;
    }


    let htmlContent = `
      <html>
        <head>
          <style>
            body { font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; color: #333; padding: 20px; }
            h1 { color: #ff9800; text-align: center; border-bottom: 2px solid #ff9800; padding-bottom: 10px; margin-bottom: 20px; }
            .info-box { text-align: center; font-size: 16px; margin-bottom: 30px; color: #555; background: #fff3e0; padding: 15px; border-radius: 8px; border: 1px solid #ffe0b2; }
            .chart-container { margin-bottom: 40px; text-align: center; page-break-inside: avoid; }
            .chart-title { font-size: 18px; font-weight: bold; margin-bottom: 15px; color: #444; text-align: left; border-left: 4px solid #4caf50; padding-left: 10px; }
            img { max-width: 100%; height: auto; border: 1px solid #ddd; border-radius: 8px; padding: 10px; background: white; box-shadow: 0 2px 5px rgba(0,0,0,0.05); }
          </style>
        </head>
        <body>
          <h1>${pdfFormalTitle}</h1>
          <div class="info-box">
            <p style="margin: 5px 0;"><strong>📆 統計區間：</strong> ${startDate} ~ ${endDate}</p>
            <p style="margin: 5px 0;"><strong>📝 報告產出時間：</strong> ${formatDateTime(new Date())}</p>
          </div>
          
          ${aiSectionHtml}
          
          <div class="chart-container">
            <div class="chart-title">📊 每日熱量與水量趨勢</div>
            <img src="${chartsBase64.intake}" />
          </div>
          
          <div class="chart-container">
            <div class="chart-title">🕒 進食時間分佈</div>
            <img src="${chartsBase64.scatter}" />
          </div>
          
          <div class="chart-container">
            <div class="chart-title">⚖️ 體重變化</div>
            <img src="${chartsBase64.weight}" />
          </div>
          
          <div class="chart-container">
            <div class="chart-title">🏥 排泄與異常狀態追蹤</div>
            <img src="${chartsBase64.poop}" />
          </div>

          <div class="chart-container">
            <div class="chart-title">💖 每日平均食慾趨勢</div>
            <img src="${chartsBase64.appetite}" />
          </div>
        </body>
      </html>
    `;

    let blob = Utilities.newBlob(htmlContent, MimeType.HTML).getAs(MimeType.PDF);
    blob.setName(`${catName}_健康報告_${startDate}.pdf`);

    MailApp.sendEmail({
      to: emailList,
      subject: emailSubject,
      htmlBody: `<p>您好！</p><p>附件是 <b>${catName}</b> 從 ${startDate} 到 ${endDate} 的健康分析報告。</p><p>此報告由您的專屬貓咪健康管家自動產生，祝您的貓咪健康快樂！🐱</p>`,
      attachments: [blob]
    });

    return { success: true };
  } catch (error) {
    throw new Error("產生或寄送 PDF 失敗：" + error.toString());
  }
}