function autoFeedPlan0300() {  
  var sheetNames = [CONFIG.SHEETS.FENDA_FEED, CONFIG.SHEETS.SEVENSEVEN_FEED];  

  for (var i = 0; i < sheetNames.length; i++) {
    eatFood({event: "凌晨"}, sheetNames[i]);
  }
}

function getSchedules() {
  let props = PropertiesService.getScriptProperties().getProperty('feedingSchedules');
  return props ? JSON.parse(props) : [];
}

function addSchedule(scheduleData) {
  let schedules = getSchedules();
  scheduleData.id = new Date().getTime().toString(); 
  schedules.push(scheduleData);
  
  PropertiesService.getScriptProperties().setProperty('feedingSchedules', JSON.stringify(schedules));
  setupMasterTrigger();
  
  return schedules;
}

function deleteSchedule(id) {
  let schedules = getSchedules();
  schedules = schedules.filter(s => s.id !== id);
  PropertiesService.getScriptProperties().setProperty('feedingSchedules', JSON.stringify(schedules));
  
  if (schedules.length === 0) {
    removeMasterTrigger();
  }
  return schedules;
}

function setupMasterTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  const hasTrigger = triggers.some(t => t.getHandlerFunction() === CONFIG.TRIGGER_HANDLER);
  
  if (!hasTrigger) {
    ScriptApp.newTrigger(CONFIG.TRIGGER_HANDLER)
      .timeBased()
      .everyHours(CONFIG.TRIGGER_INTERVAL)
      .create();
  }
}

function removeMasterTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === CONFIG.TRIGGER_HANDLER) {
      ScriptApp.deleteTrigger(t);
    }
  });
}

function runScheduledFeedings() {
  const now = getNowInTimezone();
  const tzStr = Utilities.formatDate(now, CONFIG.TIMEZONE, "yyyy-MM-dd HH:mm");
  const currentDate = tzStr.split(' ')[0];
  const currentHour = getHourFromTime(tzStr.split(' ')[1]);

  let schedules = getSchedules();
  
  schedules.forEach(schedule => {
    let schedHour = getHourFromTime(schedule.time);
    
    if (currentHour === schedHour) {
       let autoData = {
          catName: schedule.catName,
          date: currentDate,
          time: schedule.time,
          records: [{
            category: schedule.category,
            foodName: schedule.foodName,
            amount: schedule.amount,
            water: schedule.water || 0,
            notes: "🤖 自動排程寫入"
          }]
       };
       addCatRecord(autoData);
    }
  });
}