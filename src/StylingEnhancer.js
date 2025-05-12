// === StylingEnhancer.gs ===

function colorByFillRate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const sheetName = sheet.getName();
  if (!/^[A-Z][a-z]{2}\d{4}$/.test(sheetName)) return;

  const paramSheet = ss.getSheetByName("Param");
  const paramValues = paramSheet.getRange("A2:B").getValues();
  const paramMap = Object.fromEntries(paramValues.filter(r => r[0] && r[1]));

  const redThreshold = Number(paramMap["fill_rate_red"] || 80);
  const yellowThreshold = Number(paramMap["fill_rate_yellow"] || 50);

  const lastRow = sheet.getLastRow();
  const fillRates = sheet.getRange(2, 8, lastRow - 1, 1).getValues(); // H列: 埋まり率
  const currentFillColors = sheet.getRange(2, 8, lastRow - 1, 1).getBackgrounds(); // H列: 現在の背景色
  const rowBgColors = sheet.getRange(2, 2, lastRow - 1, 1).getBackgrounds(); // B列: 曜日の背景色を参照

  const colors = fillRates.map((row, i) => {
    const rate = row[0];
    const currentColor = currentFillColors[i][0];
    const bg = rowBgColors[i][0];

    // 判定：グレー（#EEEEEE）= 土日祝、水色（#B3E5FC）= 休暇
    const isNonBusinessDay = bg.toLowerCase() === "#eeeeee" || bg.toLowerCase() === "#b3e5fc";

    if (rate === "" || rate === null || typeof rate !== "number") return [currentColor];
    if (isNonBusinessDay) return [currentColor];

    if (rate >= redThreshold) return ["#f4cccc"];
    if (rate >= yellowThreshold) return ["#fff2cc"];
    return ["#d9ead3"];
  });

  sheet.getRange(2, 8, colors.length, 1).setBackgrounds(colors);
}

function updateStatsAndStyle() {
  updateStatsForCurrentSheet();
  colorByFillRate();
}
