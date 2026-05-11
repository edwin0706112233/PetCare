// ================= ✨ 共用工具函數庫 =================

/**
 * 取得 Google Sheet 實例
 * @returns {SpreadsheetApp.Spreadsheet}
 */
function getSpreadsheet() {
  return SpreadsheetApp.openById(CONFIG.SHEET_ID);
}

/**
 * 取得指定 Sheet
 * @param {string} sheetName - Sheet 名稱
 * @returns {SpreadsheetApp.Sheet|null}
 */
function getSheet(sheetName) {
  const ss = getSpreadsheet();
  return ss.getSheetByName(sheetName) || null;
}

/**
 * 根據使用者 ID 取得貓咪 Sheet 名稱
 * @param {number} userId - 使用者 ID (77 或其他)
 * @returns {string} Sheet 名稱
 */
function getCatSheetName(userId) {
  return (userId == 77) ? CONFIG.SHEETS.SEVENSEVEN_FEED : CONFIG.SHEETS.FENDA_FEED;
}

/**
 * 根據使用者 ID 取得貓咪 Sheet 名稱 (用於事件名稱)
 * @param {string} eventName - 事件名稱 (例如 "凌晨")
 * @returns {string} 對應的 Sheet 名稱
 */
function getCatSheetByEvent(eventName) {
  // 此函式用於排程中取得貓咪名稱
  return eventName ? CONFIG.SHEETS.FENDA_FEED : CONFIG.SHEETS.FENDA_FEED;
}

/**
 * 格式化日期 (YYYY/MM/DD 格式)
 * @param {Date} date - 日期物件
 * @returns {string} 格式化後的日期字串
 */
function formatDate(date) {
  if (!(date instanceof Date) || isNaN(date)) return "";
  return Utilities.formatDate(date, CONFIG.TIMEZONE, "yyyy/MM/dd");
}

/**
 * 格式化時間 (HH:mm 格式)
 * @param {Date|string} time - 時間物件或字串
 * @returns {string} 格式化後的時間字串
 */
function formatTime(time) {
  if (time instanceof Date) {
    return Utilities.formatDate(time, CONFIG.TIMEZONE, "HH:mm");
  }
  return String(time);
}

/**
 * 格式化完整日期時間 (YYYY/MM/dd HH:mm 格式)
 * @param {Date} date - 日期物件
 * @returns {string} 格式化後的日期時間字串
 */
function formatDateTime(date) {
  if (!(date instanceof Date) || isNaN(date)) return "";
  return Utilities.formatDate(date, CONFIG.TIMEZONE, "yyyy/MM/dd HH:mm");
}

/**
 * 取得簡化日期格式 (M/d 格式)
 * @param {Date} date - 日期物件
 * @returns {string} 簡化日期字串
 */
function formatSimpleDate(date) {
  if (!(date instanceof Date) || isNaN(date)) return "";
  return (date.getMonth() + 1) + "/" + date.getDate();
}

/**
 * 取得現在時間 (台北時區)
 * @returns {Date} 當前時間
 */
function getNowInTimezone() {
  return new Date();
}

/**
 * 取得今天的開始時間 (00:00:00)
 * @returns {Date}
 */
function getTodayStart() {
  const now = new Date();
  return new Date(now.getFullYear(), now.getMonth(), now.getDate());
}

/**
 * 取得今天的結束時間 (23:59:59)
 * @returns {Date}
 */
function getTodayEnd() {
  const today = getTodayStart();
  today.setHours(23, 59, 59, 999);
  return today;
}

/**
 * 取得昨天的開始時間 (00:00:00)
 * @returns {Date}
 */
function getYesterdayStart() {
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  return new Date(yesterday.getFullYear(), yesterday.getMonth(), yesterday.getDate());
}

/**
 * 驗證日期是否有效且在指定範圍內
 * @param {Date} date - 要驗證的日期
 * @param {Date} startDate - 開始日期
 * @param {Date} endDate - 結束日期
 * @returns {boolean}
 */
function isDateInRange(date, startDate, endDate) {
  if (!(date instanceof Date) || isNaN(date)) return false;
  return date >= startDate && date <= endDate;
}

/**
 * 取得貓咪的顯示名稱
 * @param {string} sheetName - Sheet 名稱
 * @returns {string} 貓咪顯示名稱
 */
function getCatDisplayName(sheetName) {
  return CONFIG.CAT_DISPLAY_NAMES[sheetName] || "貓咪";
}

/**
 * 將字串轉換為數字，如果失敗則回傳預設值
 * @param {any} value - 要轉換的值
 * @param {number} defaultValue - 預設值
 * @returns {number}
 */
function toNumber(value, defaultValue = 0) {
  const num = parseFloat(value);
  return isNaN(num) ? defaultValue : num;
}

/**
 * 安全地將字串轉換為字串，去除多餘空白
 * @param {any} value - 要轉換的值
 * @returns {string}
 */
function toString(value) {
  return String(value || "").trim();
}

/**
 * 建立空日期物件 (僅時間部分)
 * @param {number} hours - 小時
 * @param {number} minutes - 分鐘
 * @param {number} seconds - 秒數
 * @returns {Date}
 */
function createTime(hours = 0, minutes = 0, seconds = 0) {
  return new Date(0, 0, 0, hours, minutes, seconds);
}

/**
 * 建立日期物件 (僅日期部分)
 * @param {number} year - 年
 * @param {number} month - 月 (1-12)
 * @param {number} date - 日
 * @returns {Date}
 */
function createDate(year, month, date) {
  return new Date(year, month - 1, date);
}

/**
 * 比較兩個日期是否為同一天 (忽略時間)
 * @param {Date} date1 - 第一個日期
 * @param {Date} date2 - 第二個日期
 * @returns {boolean}
 */
function isSameDay(date1, date2) {
  if (!(date1 instanceof Date) || !(date2 instanceof Date)) return false;
  return date1.getFullYear() === date2.getFullYear() &&
         date1.getMonth() === date2.getMonth() &&
         date1.getDate() === date2.getDate();
}

/**
 * 取得小時部分 (從日期或時間字串)
 * @param {Date|string} dateOrTime - 日期或時間
 * @returns {number} 小時值 (0-23)
 */
function getHourFromTime(dateOrTime) {
  if (dateOrTime instanceof Date) {
    return dateOrTime.getHours();
  }
  const timeStr = String(dateOrTime).split(':')[0];
  return parseInt(timeStr, 10) || 0;
}
