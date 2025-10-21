function addTaResourceMenus_(ui) {
  (ui || SpreadsheetApp.getUi())
    .createMenu('各專業 TA 專案資源')
    .addItem('人力資源', 'consolidateHeadcount')
    .addItem('所有資源', 'consolidateAllResources')
    .addToUi();
}

function consolidateHeadcount() {
  syncTaResourceSheets();
}

function consolidateAllResources() {
  const SOURCE_SHEET_PATTERN = '1.2_各事業 TA 預算規劃_$';
  const TARGET_SPREADSHEET_ID = '1Lg520_67UD8MtVwhK1hJgybh377knWyz8iQPjwKDLbM';
  const TARGET_SHEET_NAME = '2.6_TA 專案所需資源（匯總）';

  const SOURCE_SPREADSHEET_IDS = [
    '10jcSlS4RvuGm1DK4KnyKz0YBVxHE_2OCnrBgAgqui9U', // QLR 2026 管理中心預算表
    '1X6c37n6s4XQumB4mD-N_M_Pqv4ao9f9S7Zea6c6TXRI', // QLR 2026 研發中心預算表
    '1-jof2z4-D2KbMRq2F7h0BW_p7cQy95Ppb-R1k9b6oeQ', // QLR 2026 行銷營運中心預算表
    '17a5CYVYBiJgD85EYyV4GGFW4Z8DBbc0mfNcGCyAeml4', // iCHEF 2026 管理中心預算表
    '1GPAPW3ZM3lY1Qmpb9LocnB27XWJ8AGybGhMNAy2KH0c', // iCHEF 2026 財務中心預算表
    '1mYeC6DpUqFvnce9leb-098wSYCGqfOR5_HGDM2bcbEc', // iCHEF 2026 客戶價值中心預算表
    '1Pt2ru3sBxa8TpMfTsfQLjJqdTzZZk0AYpAJvc9TjFY8', // iCHEF 2026 行銷營運中心預算表
    '1ALyHp6xEMt8xt3T9zpEA52ZUzaDPy-QT-rRh6gU75_g', // iCHEF 2026 研發中心預算表
    '1Y1KcaypEZGn_EePnDRYgj12mNw9yEvin7AZRsPhL_7E', // iCHEF 2026 策略資料中心預算表
  ];

  const filters = {
    subsidiary: 'QLR',
    businessUnit: 'OMO',
    siExclusions: ['Maintenance', 'Corporation'],
    taExclusions: ['Baseline', 'Corporation'],
    projectCodeExclusions: ['NA']
  };

  let header = null;
  const aggregatedRows = [];

  SOURCE_SPREADSHEET_IDS.forEach(spreadsheetId => {
    const dataset = loadDatasetBySheetPattern_(spreadsheetId, SOURCE_SHEET_PATTERN);
    if (!header || !header.length) header = dataset.header;
    if (!dataset.header || !dataset.header.length) return;
    const matchedRows = dataset.rows.filter(row => rowMatchesFilters_(row, filters));
    aggregatedRows.push(...matchedRows);
  });

  if (!header || !header.length) {
    throw new Error('來源試算表沒有可用的標題列');
  }

  writeToTarget_(TARGET_SPREADSHEET_ID, TARGET_SHEET_NAME, header, aggregatedRows, 2);
}

function syncTaResourceSheets() {
  const SOURCE_SHEET_NAME = '事業人力資源配置匯總表(自動彙整)';
  const TARGET_SPREADSHEET_ID = '1Lg520_67UD8MtVwhK1hJgybh377knWyz8iQPjwKDLbM';
  const TARGET_SHEET_NAME = '2.5_TA 專案所需資源（人力）';

  const SOURCE_SPREADSHEET_IDS = [
    '10jcSlS4RvuGm1DK4KnyKz0YBVxHE_2OCnrBgAgqui9U', // QLR 2026 管理中心預算表
    '1X6c37n6s4XQumB4mD-N_M_Pqv4ao9f9S7Zea6c6TXRI', // QLR 2026 研發中心預算表
    '1-jof2z4-D2KbMRq2F7h0BW_p7cQy95Ppb-R1k9b6oeQ', // QLR 2026 行銷營運中心預算表
    '17a5CYVYBiJgD85EYyV4GGFW4Z8DBbc0mfNcGCyAeml4', // iCHEF 2026 管理中心預算表
    '1GPAPW3ZM3lY1Qmpb9LocnB27XWJ8AGybGhMNAy2KH0c', // iCHEF 2026 財務中心預算表
    '1mYeC6DpUqFvnce9leb-098wSYCGqfOR5_HGDM2bcbEc', // iCHEF 2026 客戶價值中心預算表
    '1Pt2ru3sBxa8TpMfTsfQLjJqdTzZZk0AYpAJvc9TjFY8', // iCHEF 2026 行銷營運中心預算表
    '1ALyHp6xEMt8xt3T9zpEA52ZUzaDPy-QT-rRh6gU75_g'  // iCHEF 2026 研發中心預算表
  ];

  const filters = {
    subsidiary: 'QLR',
    businessUnit: 'OMO',
    siExclusions: ['Maintenance', 'Corporation'],
    taExclusions: ['Baseline', 'Corporation'],
    projectCodeExclusions: ['NA']
  };

  let header = null;
  const aggregatedRows = [];

  SOURCE_SPREADSHEET_IDS.forEach(spreadsheetId => {
    const dataset = loadSourceDataset_(spreadsheetId, SOURCE_SHEET_NAME);
    if (!header && dataset.header.length) header = dataset.header;
    if (!dataset.header.length) return;
    const matchedRows = dataset.rows.filter(row => rowMatchesFilters_(row, filters));
    aggregatedRows.push(...matchedRows);
  });

  if (!header || !header.length) {
    throw new Error('來源試算表沒有可用的標題列');
  }

  writeToTarget_(TARGET_SPREADSHEET_ID, TARGET_SHEET_NAME, header, aggregatedRows);
}

function loadDatasetBySheetPattern_(spreadsheetId, sheetNamePattern) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = findSheetByPattern_(ss, sheetNamePattern);
  if (!sheet) {
    throw new Error(`找不到來源分頁，需符合名稱模式：${sheetNamePattern}，試算表 ID：${spreadsheetId}`);
  }
  return readSheetDataset_(sheet);
}

function loadSourceDataset_(spreadsheetId, sheetName) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`找不到來源分頁：${sheetName}`);

  const dataset = readSheetDataset_(sheet);
  if (!dataset.header.length) {
    throw new Error('來源分頁沒有資料');
  }
  return dataset;
}

function findSheetByPattern_(spreadsheet, sheetNamePattern) {
  const normalizedPattern = normalizeString_(sheetNamePattern);
  const usePrefixMatch = normalizedPattern.endsWith('$');
  const effectivePattern = usePrefixMatch
    ? normalizedPattern.slice(0, -1)
    : normalizedPattern;

  return spreadsheet.getSheets().find(sheet => {
    const name = normalizeString_(sheet.getName());
    if (usePrefixMatch) {
      return name.slice(0, effectivePattern.length) === effectivePattern;
    }
    return name === effectivePattern;
  }) || null;
}

function readSheetDataset_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow < 1 || lastColumn < 1) return { header: [], rows: [] };

  const header = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const rows = lastRow > 1
    ? sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues()
    : [];

  return { header, rows };
}

function rowMatchesFilters_(row, filters) {
  const subsidiary = normalizeString_(row[0]);
  if (filters.subsidiary && subsidiary !== filters.subsidiary) return false;

  const businessUnit = normalizeString_(row[1]);
  if (filters.businessUnit && businessUnit !== filters.businessUnit) return false;

  const siId = normalizeString_(row[2]);
  if (Array.isArray(filters.siExclusions) && filters.siExclusions.includes(siId)) return false;

  const taId = normalizeString_(row[3]);
  if (Array.isArray(filters.taExclusions) && filters.taExclusions.includes(taId)) return false;

  const projectCode = normalizeString_(row[4]);
  if (Array.isArray(filters.projectCodeExclusions) && filters.projectCodeExclusions.includes(projectCode)) return false;

  return true;
}

function writeToTarget_(spreadsheetId, sheetName, header, rows, startColumn) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`找不到目標分頁：${sheetName}`);

  const effectiveStart = Math.max(1, startColumn || 1);
  const requiredRows = rows.length + 1;
  const requiredColumns = header.length
    ? effectiveStart - 1 + header.length
    : effectiveStart;

  ensureCapacity_(sheet, requiredRows, requiredColumns);
  clearTargetRange_(sheet, effectiveStart);

  if (header.length) {
    sheet.getRange(1, effectiveStart, 1, header.length).setValues([header]);
  }
  if (rows.length) {
    sheet.getRange(2, effectiveStart, rows.length, header.length).setValues(rows);
  }
}

function ensureCapacity_(sheet, requiredRows, requiredColumns) {
  const currentMaxRows = sheet.getMaxRows();
  if (currentMaxRows < requiredRows) {
    sheet.insertRowsAfter(currentMaxRows, requiredRows - currentMaxRows);
  }

  const currentMaxColumns = sheet.getMaxColumns();
  if (currentMaxColumns < requiredColumns) {
    sheet.insertColumnsAfter(currentMaxColumns, requiredColumns - currentMaxColumns);
  }
}

function clearTargetRange_(sheet, startColumn) {
  const maxRows = sheet.getMaxRows();
  const maxColumns = sheet.getMaxColumns();
  const columnCount = maxColumns - startColumn + 1;
  if (columnCount <= 0) return;
  sheet.getRange(1, startColumn, maxRows, columnCount).clearContent();
}

function normalizeString_(value) {
  if (value === null || value === undefined) return '';
  if (typeof value === 'string') return value.trim();
  return String(value).trim();
}
