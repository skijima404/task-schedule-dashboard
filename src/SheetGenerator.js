function createMonthlySheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const today = new Date();
  const sixMonthsLater = new Date(today);
  sixMonthsLater.setMonth(today.getMonth() + 6);

  const startDate = new Date(today.getFullYear(), today.getMonth(), 1);
  const endDate = new Date(sixMonthsLater.getFullYear(), sixMonthsLater.getMonth() + 1, 0);

  let date = new Date(startDate);

  while (date <= endDate) {
    const year = date.getFullYear();
    const monthIndex = date.getMonth();
    const monthName = Utilities.formatDate(date, Session.getScriptTimeZone(), "MMM"); // e.g. May
    const sheetName = `${monthName}${year}`; // e.g. May2025

    if (!ss.getSheetByName(sheetName)) {
      const sheet = ss.insertSheet(sheetName);
      Logger.log(`シート '${sheetName}' を作成しました`);

      // ヘッダー行
      const headers = [
        "日付", "曜日", "稼働時間[h]", "会議数", "会議合計時間[h]",
        "タスク可能時間[h]", "1時間枠数", "埋まり率[%]"
      ];
      sheet.appendRow(headers);

      // 該当月の最終日までループ
      const lastDay = new Date(year, monthIndex + 1, 0).getDate();

      for (let d = 1; d <= lastDay; d++) {
        const dateObj = new Date(year, monthIndex, d);
        const formattedDate = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "yyyy-MM-dd");
        const formattedDay = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "EEE");
        sheet.appendRow([formattedDate, formattedDay]);
      }

      // 数値列の小数点1桁フォーマット設定
      sheet.getRange("E2:E").setNumberFormat("0.0");
      sheet.getRange("F2:F").setNumberFormat("0.0");
      sheet.getRange("H2:H").setNumberFormat("0.0");

    } else {
      Logger.log(`シート '${sheetName}' は既に存在するためスキップしました`);
    }

    // 翌月へ
    date.setMonth(date.getMonth() + 1);
    date.setDate(1);
  }

  SpreadsheetApp.getUi().alert("6ヶ月分の月別シート作成が完了しました！");
  colorWeekendsAndHolidays();
  colorMyVacations(); // ← もし常時含めるならここに追加してもOK
}