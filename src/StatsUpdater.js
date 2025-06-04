function updateStatsForCurrentSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const sheetName = sheet.getName();
  if (!/^[A-Z][a-z]{2}\d{4}$/.test(sheetName)) {
    SpreadsheetApp.getUi().alert("月別シート（例: May2025）をアクティブにしてください。");
    return;
  }

  const timezone = Session.getScriptTimeZone();
  const calendarId = Session.getActiveUser().getEmail();
  const calendar = CalendarApp.getCalendarById(calendarId);

  const lastRow = sheet.getLastRow();
  const dates = sheet.getRange("A2:B" + lastRow).getValues(); // [日付, 曜日]

  for (let i = 0; i < dates.length; i++) {
    const [dateStr, weekday] = dates[i];

    if (weekday === "Sat" || weekday === "Sun") {
      sheet.getRange(i + 2, 3, 1, 6).setValues([["", "", "", "", "", ""]]);
      continue;
    }

    const dateObj = new Date(dateStr);
    const start = new Date(dateObj.setHours(9, 0, 0));
    const end = new Date(dateObj.setHours(18, 0, 0));

    const rawEvents = calendar.getEvents(start, end);

    const events = rawEvents.filter(e => {
      const status = e.getMyStatus?.();
      const isOwner = e.isOwnedByMe?.();
      const title = e.getTitle?.()?.toLowerCase?.() || "";

      // 🚫 除外条件：終日イベント かつ タイトルに OFF を含まないもの
      if (e.isAllDayEvent() && !title.includes("off")) return false;

      // ✅ 含める条件
      return (
        status === CalendarApp.GuestStatus.YES ||
        status === CalendarApp.GuestStatus.MAYBE ||
        status === null ||
        status === undefined ||
        status === "OWNER" ||
        isOwner === true
      );
    });

    const count = events.length;
    const merged = mergeOverlappingIntervals(events.map(e => ({
      start: e.getStartTime(),
      end: e.getEndTime()
    })));

    const totalHours = merged.reduce((sum, interval) =>
      sum + (interval.end - interval.start) / (1000 * 60 * 60), 0);

    const workingHours = 8.0;
    const available = Math.max(0, workingHours - totalHours);
    const oneHourBlocks = Math.floor(available);
    const ratio = Math.round((totalHours / workingHours) * 100);

    sheet.getRange(i + 2, 3, 1, 6).setValues([[workingHours, count, totalHours, available, oneHourBlocks, ratio]]);
  }
}

// 時間帯の重複をマージ
function mergeOverlappingIntervals(intervals) {
  if (intervals.length === 0) return [];
  intervals.sort((a, b) => a.start - b.start);
  const result = [intervals[0]];
  for (let i = 1; i < intervals.length; i++) {
    const last = result[result.length - 1];
    const current = intervals[i];
    if (current.start <= last.end) {
      last.end = new Date(Math.max(last.end, current.end));
    } else {
      result.push(current);
    }
  }
  return result;
}

function updateCalendarStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const today = new Date();
  const sheets = ss.getSheets();

  for (let offset = 0; offset < 3; offset++) {
    const targetDate = new Date(today.getFullYear(), today.getMonth() + offset, 1);
    const monthName = targetDate.toLocaleString("en-US", { month: "short" }) +
                      targetDate.getFullYear(); // "Jun2025" の形式

    const sheet = sheets.find(s => s.getName() === monthName);
    if (sheet) {
      ss.setActiveSheet(sheet);
      updateStatsAndStyle(); // StylingEnhancer.js 側の関数
    }
  }
  SpreadsheetApp.getUi().alert("直近3ヶ月のシートを更新しました！");
}