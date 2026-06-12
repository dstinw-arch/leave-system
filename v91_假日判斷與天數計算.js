// ============================================================
// v91：115年(2026)辦公日曆 — 假日判斷 + 天數自動計算
// 來源：115年辦公日曆表.xlsx（已逐格解析底色，2026 無補班六）
// ============================================================

// --- 1. 平日國定假日（週六日另由 getDay() 判斷）---
const HOLIDAYS_2026 = {
  '2026-01-01': '元旦',
  '2026-02-16': '除夕',
  '2026-02-17': '春節',
  '2026-02-18': '春節(初二)',
  '2026-02-19': '春節(初三)',
  '2026-02-20': '春節補假',
  '2026-02-27': '228補假',
  '2026-04-03': '兒童節補假',
  '2026-04-06': '清明節補假',
  '2026-05-01': '勞動節',
  '2026-06-19': '端午節',
  '2026-09-25': '中秋節',
  '2026-09-28': '教師節',
  '2026-10-09': '國慶日補假',
  '2026-10-26': '光復節補假',
  '2026-12-25': '行憲紀念日'
};

function fmtDate(d) {
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${d.getFullYear()}-${m}-${day}`;
}

// 'YYYY-MM-DD' 或 Date 皆可傳入
function isWorkday(input) {
  const d = typeof input === 'string' ? new Date(input + 'T00:00:00') : input;
  const wd = d.getDay();
  if (wd === 0 || wd === 6) return false;       // 週六日
  return !HOLIDAYS_2026[fmtDate(d)];            // 平日國定假日
}

// 回傳假日名稱（非假日回傳 null），給提示訊息用
function holidayName(dateStr) {
  const d = new Date(dateStr + 'T00:00:00');
  if (HOLIDAYS_2026[dateStr]) return HOLIDAYS_2026[dateStr];
  if (d.getDay() === 0) return '星期日';
  if (d.getDay() === 6) return '星期六';
  return null;
}

// --- 2. 天數自動計算 ---
// startTime / endTime 直接傳下拉選單的值，例如 '08:40' '12:50' '17:40'
// 規則：只計工作日；起始時間 >= 12:00 → 第一個工作日算 0.5；
//       結束時間 <= 13:00 → 最後一個工作日算 0.5
function calcLeaveDays(startDateStr, endDateStr, startTime, endTime) {
  const start = new Date(startDateStr + 'T00:00:00');
  const end = new Date(endDateStr + 'T00:00:00');
  if (end < start) return 0;

  const workdays = [];
  for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {
    if (isWorkday(d)) workdays.push(fmtDate(d));
  }
  if (workdays.length === 0) return 0;

  let days = workdays.length;
  const startHour = parseInt((startTime || '08:40').split(':')[0], 10);
  const endHour = parseInt((endTime || '17:40').split(':')[0], 10);

  if (startHour >= 12 && workdays[0] === startDateStr) days -= 0.5;   // 下午才開始
  if (endHour <= 13 && workdays[workdays.length - 1] === endDateStr) days -= 0.5; // 中午就結束
  return Math.max(days, 0);
}

// --- 3. 申請假單頁：日期/時間變更時自動帶入天數 ---
// 把四個欄位的 id 換成你實際的 id
function autoFillDays() {
  const s = document.getElementById('startDate').value;
  const e = document.getElementById('endDate').value;
  if (!s || !e) return;

  // 起訖日若選到放假日，直接提醒並擋下
  for (const [label, ds] of [['開始日期', s], ['結束日期', e]]) {
    const name = holidayName(ds);
    if (name) {
      alert(`${label} ${ds.replaceAll('-', '/')} 是「${name}」，當天本來就放假，請改選工作日`);
      return;
    }
  }

  const st = document.getElementById('startTime').value; // 例 '12:50'
  const et = document.getElementById('endTime').value;   // 例 '17:40'
  const days = calcLeaveDays(s, e, st, et);
  document.getElementById('leaveDays').value = days;     // 天數欄位
}

// 綁定（放在初始化的地方）
['startDate', 'endDate', 'startTime', 'endTime'].forEach(id => {
  const el = document.getElementById(id);
  if (el) el.addEventListener('change', autoFillDays);
});

// --- 4. 行事曆頁：假日不要畫請假底色 ---
// 在 renderCalendar() 畫每一格時，原本是：
//   if (該日在某筆請假的 start~end 範圍內) → 加底色/紅點
// 改成多一個條件：
//   if (在請假範圍內 && isWorkday(該日)) → 才加底色/紅點
// 這樣 6/13(六)、6/14(日)、6/19(端午) 就不會被標成請假日
