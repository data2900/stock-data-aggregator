function setConditionalFormattingForStockSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("StockData_統合");
  if (!sheet) throw new Error("❌ シートが見つかりません");

  const red = "#f4cccc"; // 増加（赤）
  const blue = "#cfe2f3"; // 減少（青）
  const maxRows = sheet.getLastRow();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const targetPrefixes = [
    "株価", "予想PER", "予想配当利回り", "PBR（実績）", "ROE（予想）", "株式益回り（予想）",
    "増収率", "経常増益率", "売上高経常利益率", "ROE", "ROA", "株主資本比率", "配当性向",
    "レーティング", "売上高予想", "経常利益予想", "規模", "割安度", "成長", "収益性",
    "安全性", "リスク", "リターン", "流動性", "トレンド", "為替", "テクニカル"
  ];

  const percentPrefixes = [
    "予想配当利回り", "ROE（予想）", "株式益回り（予想）",
    "増収率", "経常増益率", "売上高経常利益率",
    "ROE", "ROA", "株主資本比率", "配当性向"
  ];

  // プレフィックスごとに列をマッピング
  const prefixMap = {};
  headers.forEach((header, colIndex) => {
    for (const prefix of targetPrefixes) {
      if (header.startsWith(prefix + "_")) {
        if (!prefixMap[prefix]) prefixMap[prefix] = [];
        prefixMap[prefix].push({ header, col: colIndex + 1 }); // 1-indexed
      }
    }
  });

  // 各プレフィックスの列を日付降順で並び替え
  for (const prefix in prefixMap) {
    prefixMap[prefix].sort((a, b) => {
      const dateA = a.header.split("_")[1];
      const dateB = b.header.split("_")[1];
      return dateB.localeCompare(dateA); // 降順
    });
  }

  // %→小数変換：一括データ取得→変換→書き戻し
  const allPercentCols = percentPrefixes
    .flatMap(prefix => prefixMap[prefix] || [])
    .map(c => c.col);

  if (allPercentCols.length > 0) {
    const minCol = Math.min(...allPercentCols);
    const maxCol = Math.max(...allPercentCols);
    const range = sheet.getRange(2, minCol, maxRows - 1, maxCol - minCol + 1);
    const values = range.getValues();

    const colMap = {};
    allPercentCols.forEach(col => colMap[col - minCol] = true);

    for (let row = 0; row < values.length; row++) {
      for (let colOffset in colMap) {
        const cell = values[row][colOffset];
        if (typeof cell === "string" && cell.trim().endsWith("%")) {
          const num = parseFloat(cell.replace("%", ""));
          values[row][colOffset] = isNaN(num) ? cell : num / 100;
        }
      }
    }
    range.setValues(values);
  }

  // 条件付き書式作成（全ルール）
  const rules = [];
  for (const prefix in prefixMap) {
    const cols = prefixMap[prefix];
    for (let i = 0; i < cols.length - 1; i++) {
      const newerCol = cols[i].col;
      const olderCol = cols[i + 1].col;
      const newerA1 = columnToLetter(newerCol);
      const olderA1 = columnToLetter(olderCol);
      const range = sheet.getRange(`${newerA1}2:${newerA1}${maxRows}`);

      // 増加 → 赤
      rules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied(`=AND(ISNUMBER(${newerA1}2), ISNUMBER(${olderA1}2), ${newerA1}2 > ${olderA1}2)`)
          .setBackground(red)
          .setRanges([range])
          .build()
      );

      // 減少 → 青
      rules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied(`=AND(ISNUMBER(${newerA1}2), ISNUMBER(${olderA1}2), ${newerA1}2 < ${olderA1}2)`)
          .setBackground(blue)
          .setRanges([range])
          .build()
      );
    }
  }

  // 数値フォーマットの一括設定
  const allCols = Object.values(prefixMap).flat().map(c => c.col);
  const uniqueCols = [...new Set(allCols)];
  const minDataCol = Math.min(...uniqueCols);
  const maxDataCol = Math.max(...uniqueCols);
  const formatRange = sheet.getRange(2, minDataCol, maxRows - 1, maxDataCol - minDataCol + 1);
  formatRange.setNumberFormat("0.##");

  // 書式ルール設定
  sheet.clearConditionalFormatRules();
  sheet.setConditionalFormatRules(rules);

  Logger.log(`✅ 条件付き書式 ${rules.length} 件設定 & %→小数変換 & 数値フォーマット完了`);
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
