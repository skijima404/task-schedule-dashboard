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
  ui.createMenu("📊 ダッシュボード")
    .addItem("📆 今開いている月を更新", "updateStatsAndStyle")
    .addItem("🔁 全シート統計を更新", "updateCalendarStats")
    .addSeparator()
    .addItem("🎨 休暇日を反映", "colorMyVacations")
    .addItem("🌈 埋まり率で色分け", "colorByFillRate")
    .addSeparator()
    .addItem("📅 月別シートを作成", "createMonthlySheets")
    .addToUi();
}