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

// Update 3 Month Stats And Style のスクリプトはStatsUpdater.js にあります


// === 3Month Heatmap ===

function colorFillRateInHeatmap() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const heatmapSheet = ss.getSheetByName("Heatmap");
  const paramSheet = ss.getSheetByName("Param");
  const paramValues = paramSheet.getRange("A2:B").getValues();
  const paramMap = Object.fromEntries(paramValues.filter(r => r[0] && r[1]));

  const redThreshold = Number(paramMap["fill_rate_red"] || 80);
  const yellowThreshold = Number(paramMap["fill_rate_yellow"] || 50);

  const monthSheets = ss.getSheets().filter(s => /^[A-Z][a-z]{2}\d{4}$/.test(s.getName()));
  const fillRateMap = new Map();

  for (const sheet of monthSheets) {
    const data = sheet.getDataRange().getValues();
    const dateIdx = data[0].indexOf("日付");
    const rateIdx = data[0].findIndex(h => typeof h === "string" && h.match(/埋まり率.*%/));
    if (dateIdx === -1 || rateIdx === -1) continue;

    for (let i = 1; i < data.length; i++) {
      const dateStr = data[i][dateIdx];
      const rate = data[i][rateIdx];
      if (dateStr && typeof rate === "number") {
        fillRateMap.set(Utilities.formatDate(new Date(dateStr), Session.getScriptTimeZone(), "yyyy-MM-dd"), rate);
      }
    }
  }

  Logger.log("📊 FillRateMap size: " + fillRateMap.size); //

  const range = heatmapSheet.getDataRange();
  const values = range.getValues();
  const backgrounds = range.getBackgrounds();

  for (let i = 0; i < values.length; i++) {
    for (let j = 0; j < values[i].length; j++) {
      const cellVal = values[i][j];
      const cellBg = backgrounds[i][j];

      if (!(cellVal instanceof Date)) continue;

      const dateStr = Utilities.formatDate(cellVal, Session.getScriptTimeZone(), "yyyy-MM-dd");
      const rate = fillRateMap.get(dateStr);

      Logger.log("Checking: " + dateStr + " → " + rate);

      if (rate === undefined) continue;

      // すでに背景色が土日祝や休暇の色であれば上書きしない
      const bg = cellBg.toLowerCase();  // ← 小文字に変換して比較
      if (bg === "#eeeeee" || bg === "#b3e5fc") continue;

      if (rate >= redThreshold) {
        backgrounds[i][j] = "#f4cccc";
      } else if (rate >= yellowThreshold) {
        backgrounds[i][j] = "#fff2cc";
      } else {
        backgrounds[i][j] = "#d9ead3";
      }
    }
  }

  range.setBackgrounds(backgrounds);
}

function highlightTodayInHeatmap() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Heatmap");
  if (!sheet) return;

  const today = new Date();
  const todayStr = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM-dd");
  const range = sheet.getDataRange();
  const values = range.getValues();
  const borders = range.getBorder();

  for (let i = 0; i < values.length; i++) {
    for (let j = 0; j < values[i].length; j++) {
      const cell = values[i][j];
      if (Object.prototype.toString.call(cell) === "[object Date]") {
        const cellStr = Utilities.formatDate(cell, Session.getScriptTimeZone(), "yyyy-MM-dd");
        if (cellStr === todayStr) {
          sheet.getRange(i + 1, j + 1).setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
          return;
        }
      }
    }
  }
} 