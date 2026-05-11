function autoFeedPlan0300() {  
  var sheetNames = ["芬達吃飯","77吃飯"];  

  for (var i = 0; i < sheetNames.length; i++) {
    eatFood({event: "凌晨"}, sheetNames[i]);
  }
}

// ================= ✨ UI 定時排程與自動化系統 =================

// 1. 取得目前所有的排程
function getSchedules() {
  let props = PropertiesService.getScriptProperties().getProperty('feedingSchedules');
  return props ? JSON.parse(props) : [];
}

// 2. 新增排程 (由 UI 呼叫)
function addSchedule(scheduleData) {
  let schedules = getSchedules();
  // 產生唯一 ID
  scheduleData.id = new Date().getTime().toString(); 
  schedules.push(scheduleData);
  
  // 存入隱藏的系統資料庫
  PropertiesService.getScriptProperties().setProperty('feedingSchedules', JSON.stringify(schedules));
  
  // ✨ 自動檢查並建立「巡邏隊長」觸發條件
  setupMasterTrigger();
  
  return schedules;
}

// 3. 刪除排程 (由 UI 呼叫)
function deleteSchedule(id) {
  let schedules = getSchedules();
  schedules = schedules.filter(s => s.id !== id);
  PropertiesService.getScriptProperties().setProperty('feedingSchedules', JSON.stringify(schedules));
  
  // 如果排程全空了，就把觸發條件刪除以節省資源
  if (schedules.length === 0) {
    removeMasterTrigger();
  }
  return schedules;
}

// 4. 建立巡邏隊長 (每小時執行一次)
function setupMasterTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  const hasTrigger = triggers.some(t => t.getHandlerFunction() === 'runScheduledFeedings');
  
  if (!hasTrigger) {
    ScriptApp.newTrigger('runScheduledFeedings')
      .timeBased()
      .everyHours(1) // 每小時巡邏一次
      .create();
  }
}

// 5. 移除巡邏隊長
function removeMasterTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'runScheduledFeedings') {
      ScriptApp.deleteTrigger(t);
    }
  });
}

// 6. 🚨 核心執行程式 (時間到了會由 Google 伺服器在背景自動呼叫)
function runScheduledFeedings() {
  // 設定為台北時區
  const now = new Date();
  const tzStr = Utilities.formatDate(now, "Asia/Taipei", "yyyy-MM-dd HH:mm");
  const currentDate = tzStr.split(' ')[0]; // 取得 YYYY-MM-DD
  const currentHour = parseInt(tzStr.split(' ')[1].split(':')[0], 10); // 取得現在是幾點 (0~23)

  let schedules = getSchedules();
  
  schedules.forEach(schedule => {
    // schedule.time 格式為 "16:00"
    let schedHour = parseInt(schedule.time.split(':')[0], 10);
    
    // 如果「現在的小時」等於「設定的小時」，就執行餵食紀錄！
    if (currentHour === schedHour) {
       let autoData = {
          catName: schedule.catName,
          date: currentDate,
          time: schedule.time, // 寫入精準的設定時間
          records: [{
            category: schedule.category,
            foodName: schedule.foodName,
            amount: schedule.amount,
            water: schedule.water || 0,
            notes: "🤖 自動排程寫入"
          }]
       };
       // 呼叫原本寫好的寫入函數
       addCatRecord(autoData);
    }
  });
}