function setConditionalFormattingForStockSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("StockData");
  if (!sheet) throw new Error("❌ シートが見つかりません");

  const red = "#f4cccc"; // 増加（赤）
  const blue = "#cfe2f3"; // 減少（青）
  const maxRows = sheet.getLastRow();

  // [比較対象列, 比較元列]（全週分、2列ずらし済み）
  const comparePairs = [
    // 最新 vs 2週目
    ["D", "X"], ["E", "Y"], ["F", "Z"], ["G", "AA"], ["H", "AB"],
    ["I", "AC"], ["J", "AD"], ["K", "AE"], ["L", "AF"], ["M", "AG"],
    ["N", "AH"], ["O", "AI"], ["P", "AJ"], ["Q", "AK"], ["R", "AL"],
    ["S", "AM"], ["T", "AN"], ["U", "AO"], ["V", "AP"], ["W", "AQ"],

    // 2週目 vs 3週目
    ["X", "AR"], ["Y", "AS"], ["Z", "AT"], ["AA", "AU"], ["AB", "AV"],
    ["AC", "AW"], ["AD", "AX"], ["AE", "AY"], ["AF", "AZ"], ["AG", "BA"],
    ["AH", "BB"], ["AI", "BC"], ["AJ", "BD"], ["AK", "BE"], ["AL", "BF"],
    ["AM", "BG"], ["AN", "BH"], ["AO", "BI"], ["AP", "BJ"], ["AQ", "BK"],

    // 3週目 vs 4週目
    ["AR", "BL"], ["AS", "BM"], ["AT", "BN"], ["AU", "BO"], ["AV", "BP"],
    ["AW", "BQ"], ["AX", "BR"], ["AY", "BS"], ["AZ", "BT"], ["BA", "BU"],
    ["BB", "BV"], ["BC", "BW"], ["BD", "BX"], ["BE", "BY"], ["BF", "BZ"],
    ["BG", "CA"], ["BH", "CB"], ["BI", "CC"], ["BJ", "CD"], ["BK", "CE"]
  ];

  // ✅ 数値表示形式の適用（.00 を非表示にする）
  const uniqueCols = [...new Set(comparePairs.flat())];
  uniqueCols.forEach(col => {
    const range = sheet.getRange(`${col}2:${col}${maxRows}`);
    range.setNumberFormat("0.##"); // 小数点 .00 は表示しない、小数あれば表示
  });

  // ✅ 条件付き書式ルール設定
  const rules = [];
  sheet.clearConditionalFormatRules();

  comparePairs.forEach(([currentCol, prevCol]) => {
    const range = sheet.getRange(`${currentCol}2:${currentCol}${maxRows}`);

    // 増加 → 赤
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=AND(ISNUMBER(${currentCol}2), ISNUMBER(${prevCol}2), ${currentCol}2 > ${prevCol}2)`)
        .setBackground(red)
        .setRanges([range])
        .build()
    );

    // 減少 → 青
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=AND(ISNUMBER(${currentCol}2), ISNUMBER(${prevCol}2), ${currentCol}2 < ${prevCol}2)`)
        .setBackground(blue)
        .setRanges([range])
        .build()
    );
  });

  sheet.setConditionalFormatRules(rules);
  Logger.log("✅ 条件付き書式と数値表示形式（小数点以下 .00 非表示）を設定しました");
}