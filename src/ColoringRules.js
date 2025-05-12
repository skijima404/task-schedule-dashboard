function colorWeekendsAndHolidays() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const calendarId = 'en.japanese#holiday@group.v.calendar.google.com';
  const holidayCalendar = CalendarApp.getCalendarById(calendarId);
  const timezone = Session.getScriptTimeZone();

  const sheets = ss.getSheets().filter(s => /^[A-Z][a-z]{2}\d{4}$/.test(s.getName())); // e.g. May2025

  sheets.forEach(sheet => {
    const lastRow = sheet.getLastRow();
    const data = sheet.getRange(`A2:B${lastRow}`).getValues(); // 日付 + 曜日
    const dates = data.map(row => new Date(row[0]));

    // 該当月の祝日イベント取得
    const start = dates[0];
    const end = dates[dates.length - 1];
    const holidayEvents = holidayCalendar.getEvents(start, end);
    const holidays = new Set(holidayEvents.map(e =>
      Utilities.formatDate(e.getStartTime(), timezone, "yyyy-MM-dd")));

    for (let i = 0; i < data.length; i++) {
      const [dateStr, weekday] = data[i];
      const dateFormatted = Utilities.formatDate(new Date(dateStr), timezone, "yyyy-MM-dd");

      const isWeekend = weekday === "Sat" || weekday === "Sun";
      const isHoliday = holidays.has(dateFormatted);

      if (isWeekend || isHoliday) {
        sheet.getRange(i + 2, 1, 1, 8).setBackground("#EEEEEE"); // A〜H列をグレーに
      }
    }
  });

  // SpreadsheetApp.getUi().alert("土日・祝日（Google）をグレーに塗りました！");
}

function colorMyVacations() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const timezone = Session.getScriptTimeZone();
  const vacationSheet = ss.getSheetByName("MyVacation");
  if (!vacationSheet) {
    SpreadsheetApp.getUi().alert("MyVacation シートが見つかりませんでした。");
    return;
  }

  // A列の日付をセットに
  const vacations = new Set(
    vacationSheet.getRange("A2:A" + vacationSheet.getLastRow())
      .getValues()
      .flat()
      .filter(v => v)
      .map(v => Utilities.formatDate(new Date(v), timezone, "yyyy-MM-dd"))
  );

  // 月次シート（May2025など）のみに適用
  const sheets = ss.getSheets().filter(s => /^[A-Z][a-z]{2}\d{4}$/.test(s.getName()));

  sheets.forEach(sheet => {
    const lastRow = sheet.getLastRow();
    const dates = sheet.getRange(`A2:A${lastRow}`).getValues().flat();

    for (let i = 0; i < dates.length; i++) {
      const dateStr = Utilities.formatDate(new Date(dates[i]), timezone, "yyyy-MM-dd");
      if (vacations.has(dateStr)) {
        sheet.getRange(i + 2, 1, 1, 8).setBackground("#B3E5FC"); // 薄い水色など
      }
    }
  });

  // SpreadsheetApp.getUi().alert("MyVacation に登録された休暇日を塗りました！");
}


// === 3Month Heatmap ===

function colorWeekendsAndHolidaysInRange(sheet, rowStart, rowEnd) {
  const cal = CalendarApp.getCalendarById("en.japanese#holiday@group.v.calendar.google.com");
  const start = new Date();
  const end = new Date(start.getFullYear(), start.getMonth() + 3, 0);
  const holidays = cal.getEvents(start, end).map(e => Utilities.formatDate(e.getStartTime(), Session.getScriptTimeZone(), "yyyy-MM-dd"));

  let raw = sheet.getRange(rowStart - 2, 1).getValue();
  if (raw instanceof Date) {
    raw = Utilities.formatDate(raw, Session.getScriptTimeZone(), "MMMM yyyy");
  }
  const [monthNameStr, yearStr] = String(raw).split(" ");
  const monthNames = ["January", "February", "March", "April", "May", "June", "July",
                      "August", "September", "October", "November", "December"];
  const expectedMonth = monthNames.indexOf(monthNameStr);
  const expectedYear = Number(yearStr);

  const range = sheet.getRange(rowStart, 1, rowEnd - rowStart + 1, 7);
  const values = range.getValues();
  const backgrounds = range.getBackgrounds();

  for (let i = 0; i < values.length; i++) {
    for (let j = 0; j < values[i].length; j++) {
      const cell = values[i][j];
      if (Object.prototype.toString.call(cell) === "[object Date]") {
        if (cell.getMonth() !== expectedMonth || cell.getFullYear() !== expectedYear) continue;

        const day = cell.getDay();
        const dateStr = Utilities.formatDate(cell, Session.getScriptTimeZone(), "yyyy-MM-dd");
        if (day === 0 || day === 6 || holidays.includes(dateStr)) {
          backgrounds[i][j] = "#EEEEEE";
        }
      }
    }
  }

  range.setBackgrounds(backgrounds);
}

function colorMyVacationsForHeatmap() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Heatmap");
  const vacationSheet = ss.getSheetByName("MyVacation");
  if (!sheet || !vacationSheet) return;

  const vacationDates = vacationSheet.getRange("A2:A").getValues().flat().filter(v => v instanceof Date);
  const vacationStrSet = new Set(vacationDates.map(d => Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd")));

  const range = sheet.getDataRange();
  const values = range.getValues();
  const backgrounds = range.getBackgrounds();

  for (let i = 2; i < values.length; i++) {
    for (let j = 0; j < values[i].length; j++) {
      const cell = values[i][j];
      if (Object.prototype.toString.call(cell) === "[object Date]") {
        const dateStr = Utilities.formatDate(cell, Session.getScriptTimeZone(), "yyyy-MM-dd");
        if (vacationStrSet.has(dateStr)) {
          backgrounds[i][j] = "#B3E5FC";
        }
      }
    }
  }

  range.setBackgrounds(backgrounds);
} 

// Vacationようハンドラ

function colorMyVacationsEverywhere() {
  colorMyVacations();
  colorMyVacationsForHeatmap();
}