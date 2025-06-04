function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  if (sheet.getName() !== "MyVacation") return;

  const editedRange = e.range;
  const newValue = editedRange.getValue();

  // 空白（削除）されたときは何もしない
  if (!newValue) return;

  // 値が日付である場合のみ処理（保険）
  if (Object.prototype.toString.call(newValue) !== "[object Date]") return;

  // 休暇色塗り実行
  colorMyVacations();
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Task Dashboard")
    .addItem("📅 Generate Monthly Sheets", "createMonthlySheets")
    .addItem("📊 Update Stats & Style", "updateStatsAndStyle")
    .addItem("🔁 Update 3 Month Sheet Stats", "updateCalendarStats") // ★ 追加
    .addItem("🌈 Apply Fill Rate to Heatmap", "colorFillRateInHeatmap")
    .addItem("🏖️ Highlight Vacations", "colorMyVacationsEverywhere")
    .addItem("⭐ Highlight Today", "highlightTodayInHeatmap")
    .addItem("🗓️ Generate 3-Month Heatmap", "generateThreeMonthCalendar")
    .addToUi();
}