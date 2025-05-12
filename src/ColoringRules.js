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

  SpreadsheetApp.getUi().alert("土日・祝日（Google）をグレーに塗りました！");
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

  SpreadsheetApp.getUi().alert("MyVacation に登録された休暇日を塗りました！");
}