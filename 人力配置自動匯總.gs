function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('人力配置自動匯總')
    .addItem('執行匯總', 'consolidateHeadcount')
    .addToUi();
}

function onInstall() {
  onOpen();
}

function consolidateHeadcount() {
  const sourceSheetNames = [
    '1.1_各事業 TA 預算規劃_策略資料中心_人力時間配置',
    '3.1_集團共用資源_策略資料中心_人力時間配置'
  ];
  const targetSheetName = '事業人力資源配置匯總表(中心主管填)';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = ss.getSheetByName(targetSheetName);
  if (!targetSheet) throw new Error(`找不到目標分頁：${targetSheetName}`);

  const sourceSheets = sourceSheetNames.map(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) throw new Error(`找不到來源分頁：${name}`);
    return sheet;
  });

  const rawHeaderRow = sourceSheets[0]
    .getRange(1, 1, 1, sourceSheets[0].getLastColumn())
    .getValues()[0];
  const firstColumnLabelIndex = rawHeaderRow.findIndex(value =>
    typeof value === 'string' && /^Column\s+\d+$/i.test(value)
  );
  const effectiveLastCol = firstColumnLabelIndex > -1
    ? firstColumnLabelIndex
    : rawHeaderRow.length;

  const header = [rawHeaderRow.slice(0, effectiveLastCol)];
  targetSheet.clearContents();
  targetSheet.getRange(1, 1, 1, effectiveLastCol).setValues(header);

  const metricsStartIndex = (() => {
    const idx = header[0].findIndex(value => {
      if (value instanceof Date) return true;
      if (typeof value === 'string') return /^\d{4}\/?\d{1,2}/.test(value);
      return false;
    });
    return idx >= 0 ? idx : 9;
  })();

  const summaryRows = [];
  sourceSheets.forEach(sheet => {
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return;
    const rows = sheet.getRange(2, 1, lastRow - 1, effectiveLastCol).getValues();
    rows.forEach(row => {
      const hasNumbers = row.slice(metricsStartIndex)
        .some(value => typeof value === 'number' && !isNaN(value) && value !== 0);
      const hasKeys = row.slice(0, metricsStartIndex)
        .some(value => value !== '' && value !== null);
      if (hasNumbers && hasKeys) summaryRows.push(row);
    });
  });

  if (summaryRows.length) {
    targetSheet.getRange(2, 1, summaryRows.length, effectiveLastCol)
      .setValues(summaryRows);
  }
}
