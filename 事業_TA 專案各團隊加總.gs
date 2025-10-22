function aggregateTaProjectTotals() {
  const SPREADSHEET_ID = '1Lg520_67UD8MtVwhK1hJgybh377knWyz8iQPjwKDLbM';
  const SOURCE_SHEET_NAME = '2.6_TA 專案所需資源（匯總）';
  const TARGET_SHEET_NAME = '2.4_TA 專案選擇_(加總 2.6)';
  const OUTPUT_START_ROW = 24;
  const OUTPUT_START_COLUMN = 4; // column D
  const OUTPUT_HEADERS = [
    '子公司',
    '事業單位',
    'Si 編號',
    'TA 編號',
    '價值鏈專案預算代號',
    '費用單位',
    '人事費用',
    '非人事費用'
  ];

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sourceSheet = ss.getSheetByName(SOURCE_SHEET_NAME);
  const targetSheet = ss.getSheetByName(TARGET_SHEET_NAME);
  if (!sourceSheet) throw new Error(`找不到來源分頁：${SOURCE_SHEET_NAME}`);
  if (!targetSheet) throw new Error(`找不到目標分頁：${TARGET_SHEET_NAME}`);

  const dataset = readSheetDataset_(sourceSheet);
  if (!dataset.header.length) throw new Error('來源分頁沒有標題列');

  const records = buildTaProjectTotals_(dataset.header, dataset.rows);
  const requiredRows = OUTPUT_START_ROW + records.length;
  const requiredColumns = OUTPUT_START_COLUMN + OUTPUT_HEADERS.length - 1;

  ensureSheetCapacity_(targetSheet, requiredRows, requiredColumns);
  clearOutputArea_(targetSheet, OUTPUT_START_ROW, OUTPUT_START_COLUMN, OUTPUT_HEADERS.length);

  targetSheet
    .getRange(OUTPUT_START_ROW, OUTPUT_START_COLUMN, 1, OUTPUT_HEADERS.length)
    .setValues([OUTPUT_HEADERS]);

  if (records.length) {
    targetSheet
      .getRange(OUTPUT_START_ROW + 1, OUTPUT_START_COLUMN, records.length, OUTPUT_HEADERS.length)
      .setValues(records);
  }
}

function buildTaProjectTotals_(header, rows) {
  const normalize = typeof normalizeString_ === 'function'
    ? normalizeString_
    : function(value) {
        if (value === null || value === undefined) return '';
        if (typeof value === 'string') return value.trim();
        return String(value).trim();
      };

  const columnIndex = {
    subsidiary: requireColumnIndex_(header, '子公司'),
    businessUnit: requireColumnIndex_(header, '事業單位'),
    siId: requireColumnIndex_(header, 'Si 編號'),
    taId: requireColumnIndex_(header, 'TA 編號'),
    projectCode: requireColumnIndex_(header, '價值鏈專案預算代號'),
    costUnit: requireColumnIndex_(header, '費用單位'),
    costType: requireColumnIndex_(header, '費用類型'),
    note: requireColumnIndex_(header, '備註')
  };

  const yearIndex = header.indexOf('2026 year');
  const firstAmountIndex = columnIndex.note + 1;

  const totals = new Map();

  rows.forEach(row => {
    const taId = normalize(row[columnIndex.taId]);
    const projectCode = normalize(row[columnIndex.projectCode]);
    const costUnit = normalize(row[columnIndex.costUnit]);

    if (!taId && !projectCode && !costUnit) return;

    const key = [taId, projectCode, costUnit].join('||');
    if (!totals.has(key)) {
      totals.set(key, {
        subsidiary: normalize(row[columnIndex.subsidiary]),
        businessUnit: normalize(row[columnIndex.businessUnit]),
        siId: normalize(row[columnIndex.siId]),
        taId,
        projectCode,
        costUnit,
        personnel: 0,
        nonPersonnel: 0
      });
    }

    const amount = resolveRowTotal_(row, yearIndex, firstAmountIndex);
    const costType = normalize(row[columnIndex.costType]);
    const record = totals.get(key);

    if (costType === '人事費用') {
      record.personnel += amount;
    } else if (costType === '非人事費用') {
      record.nonPersonnel += amount;
    } else {
      record.nonPersonnel += amount;
    }
  });

  const sorted = Array.from(totals.values()).sort((a, b) => {
    return (
      localeCompareSafe_(a.taId, b.taId) ||
      localeCompareSafe_(a.projectCode, b.projectCode) ||
      localeCompareSafe_(a.costUnit, b.costUnit)
    );
  });

  return sorted.map(record => [
    record.subsidiary,
    record.businessUnit,
    record.siId,
    record.taId,
    record.projectCode,
    record.costUnit,
    sanitizeTotal_(record.personnel),
    sanitizeTotal_(record.nonPersonnel)
  ]);
}

function resolveRowTotal_(row, yearIndex, firstAmountIndex) {
  if (yearIndex !== -1) {
    const yearValue = parseNumber_(row[yearIndex]);
    if (!isNaN(yearValue)) return yearValue;
  }

  let sum = 0;
  for (let i = firstAmountIndex; i < row.length; i += 1) {
    if (i === yearIndex) continue;
    const value = parseNumber_(row[i]);
    if (!isNaN(value)) sum += value;
  }
  return sum;
}

function requireColumnIndex_(header, columnName) {
  const index = header.indexOf(columnName);
  if (index === -1) {
    throw new Error(`來源分頁缺少必要欄位：${columnName}`);
  }
  return index;
}

function parseNumber_(value) {
  if (typeof value === 'number') return value;
  if (typeof value === 'string') {
    const normalized = value.replace(/[,\\s]/g, '');
    if (!normalized) return NaN;
    const parsed = Number(normalized);
    return Number.isFinite(parsed) ? parsed : NaN;
  }
  return NaN;
}

function sanitizeTotal_(value) {
  const normalized = Number(value) || 0;
  const rounded = Math.round(normalized * 100) / 100;
  return Math.abs(rounded) === 0 ? 0 : rounded;
}

function ensureSheetCapacity_(sheet, requiredRows, requiredColumns) {
  if (sheet.getMaxRows() < requiredRows) {
    sheet.insertRowsAfter(sheet.getMaxRows(), requiredRows - sheet.getMaxRows());
  }
  if (sheet.getMaxColumns() < requiredColumns) {
    sheet.insertColumnsAfter(
      sheet.getMaxColumns(),
      requiredColumns - sheet.getMaxColumns()
    );
  }
}

function clearOutputArea_(sheet, startRow, startColumn, columnCount) {
  const maxRows = sheet.getMaxRows();
  if (startRow > maxRows) return;
  const rows = maxRows - startRow + 1;
  sheet.getRange(startRow, startColumn, rows, columnCount).clearContent();
}

function localeCompareSafe_(a, b) {
  const left = a || '';
  const right = b || '';
  return left.localeCompare(right, 'zh-Hant');
}
