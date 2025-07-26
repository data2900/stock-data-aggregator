function setConditionalFormattingForStockSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("StockData");
  if (!sheet) throw new Error("❌ シートが見つかりません");

  const red = "#f4cccc"; // 増加（赤）
  const blue = "#cfe2f3"; // 減少（青）
  const maxRows = sheet.getLastRow();

  // ✅ 列グループ（各週ごとの列）
  const weeklyCols = [
    getColumnRange("D", "AD"),  // 最新
    getColumnRange("AE", "BE"), // 2週目
    getColumnRange("BF", "CF"), // 3週目
    getColumnRange("CG", "DG")  // 4週目
  ];

  const comparePairs = [];

  for (let i = 0; i < weeklyCols.length - 1; i++) {
    const newerCols = weeklyCols[i];
    const olderCols = weeklyCols[i + 1];

    for (let j = 0; j < newerCols.length; j++) {
      comparePairs.push([newerCols[j], olderCols[j]]);
    }
  }

  // ✅ 数値フォーマットを設定（重複除去）
  const uniqueCols = [...new Set(comparePairs.flat())];
  uniqueCols.forEach(col => {
    const range = sheet.getRange(`${col}2:${col}${maxRows}`);
    range.setNumberFormat("0.##"); // 小数点以下非表示（必要があれば表示）
  });

  // ✅ 条件付き書式の設定
  sheet.clearConditionalFormatRules();
  const rules = [];

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

// ✨ 文字列で列範囲（例："D", "AD"）を指定し、A〜ZZ列名の配列を返す
function getColumnRange(startCol, endCol) {
  const startIndex = letterToColumn(startCol);
  const endIndex = letterToColumn(endCol);
  const cols = [];
  for (let i = startIndex; i <= endIndex; i++) {
    cols.push(columnToLetter(i));
  }
  return cols;
}

// 🔤 列番号 → A〜ZZ形式
function columnToLetter(column) {
  let temp = '';
  while (column > 0) {
    let modulo = (column - 1) % 26;
    temp = String.fromCharCode(65 + modulo) + temp;
    column = Math.floor((column - modulo) / 26);
  }
  return temp;
}

// 🔢 A〜ZZ形式 → 列番号
function letterToColumn(letter) {
  let column = 0;
  for (let i = 0; i < letter.length; i++) {
    column *= 26;
    column += letter.charCodeAt(i) - 64;
  }
  return column;
}
