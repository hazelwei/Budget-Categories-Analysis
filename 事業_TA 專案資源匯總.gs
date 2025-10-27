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
    '1Nk_TXUKf5uX0P-fmbP9jlSio6qB1pH1V4XFFsUWWIzM', // iCHEF 2026 執行長室預算表
    '1XCORsakXwHErz2AxHTTBY1a1yxsN78scGKOd7_uIPss', // iCHEF 2026 執行長 OMO 預算表
  ];

  const filters = {
    // TODO: 使用前請更新對應的 subsidiary,businessUnit。
    subsidiary: 'QLR',
    businessUnit: ['OMO', '2C'],
    siExclusions: ['Maintenance', 'Corporation'],
    taExclusions: ['Baseline', 'Corporation'],
    projectCodeExclusions: ['NA']
  };

  let header = null;
  const aggregatedRows = [];

  SOURCE_SPREADSHEET_IDS.forEach(spreadsheetId => {
    const dataset = loadDatasetBySheetPattern_(spreadsheetId, SOURCE_SHEET_PATTERN);
    if (!dataset.header || !dataset.header.length) return;

    if (!header || !header.length) header = dataset.header;

    const headerMap = buildHeaderMap_(dataset.header);
    const matchedRows = dataset.rows.filter(row => rowMatchesFilters_(row, filters, headerMap));
    aggregatedRows.push(...matchedRows);
  });

  if (!header || !header.length) {
    throw new Error('來源試算表沒有可用的標題列');
  }

  const targetSpreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  writeRowsAlignedToTarget_(targetSpreadsheetId, TARGET_SHEET_NAME, header, aggregatedRows);
}

function syncTaResourceSheets() {
  const SOURCE_SHEET_NAME = '事業人力資源配置匯總表(自動彙整)';
  const TARGET_SHEET_NAME = '2.5_TA 專案所需資源（人力）';

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
    '1Nk_TXUKf5uX0P-fmbP9jlSio6qB1pH1V4XFFsUWWIzM', // iCHEF 2026 執行長室預算表
    '1XCORsakXwHErz2AxHTTBY1a1yxsN78scGKOd7_uIPss', // iCHEF 2026 執行長 OMO 預算表
  ];

  const filters = {
    // TODO: 使用前請更新對應的 subsidiary,businessUnit。
    subsidiary: 'QLR',
    businessUnit: ['OMO', '2C'],
    siExclusions: ['Maintenance', 'Corporation'],
    taExclusions: ['Baseline', 'Corporation'],
    projectCodeExclusions: ['NA']
  };

  let header = null;
  const aggregatedRows = [];

  SOURCE_SPREADSHEET_IDS.forEach(spreadsheetId => {
    const dataset = loadSourceDataset_(spreadsheetId, SOURCE_SHEET_NAME);
    if (!dataset.header.length) return;

    if (!header) {
      header = dataset.header.slice();
    } else if (dataset.header.length > header.length) {
      const headerGrowth = dataset.header.slice(header.length);
      header.push(...headerGrowth);
      aggregatedRows.forEach(row => {
        while (row.length < header.length) row.push('');
      });
    }

    const headerMap = buildHeaderMap_(dataset.header);
    const matchedRows = dataset.rows.filter(row => rowMatchesFilters_(row, filters, headerMap));
    matchedRows.forEach(row => {
      aggregatedRows.push(normalizeRowLength_(row, header.length));
    });
  });

  if (!header || !header.length) {
    throw new Error('來源試算表沒有可用的標題列');
  }

  const targetSpreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  writeToTarget_(targetSpreadsheetId, TARGET_SHEET_NAME, header, aggregatedRows);
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

function rowMatchesFilters_(row, filters, headerMap) {
  const resolver = createRowValueResolver_(row, headerMap);

  const subsidiary = normalizeString_(resolver('子公司', 0));
  if (!matchesFilter_(subsidiary, filters.subsidiary)) return false;

  const businessUnit = normalizeString_(resolver('事業單位', 1));
  if (!matchesFilter_(businessUnit, filters.businessUnit)) return false;

  const siId = normalizeString_(resolver('Si 編號', 2));
  if (Array.isArray(filters.siExclusions) && filters.siExclusions.includes(siId)) return false;

  const taId = normalizeString_(resolver('TA 編號', 3));
  if (Array.isArray(filters.taExclusions) && filters.taExclusions.includes(taId)) return false;

  const projectCode = normalizeString_(resolver('價值鏈專案預算代號', 4));
  if (Array.isArray(filters.projectCodeExclusions) && filters.projectCodeExclusions.includes(projectCode)) return false;

  return true;
}

function createRowValueResolver_(row, headerMap) {
  const effectiveMap = headerMap || null;

  return function(columnName, fallbackIndex) {
    if (effectiveMap && columnName) {
      const key = normalizeString_(columnName);
      if (key && Object.prototype.hasOwnProperty.call(effectiveMap, key)) {
        const index = effectiveMap[key];
        if (typeof index === 'number' && index >= 0 && index < row.length) {
          return row[index];
        }
      }
    }
    if (typeof fallbackIndex === 'number' && fallbackIndex >= 0 && fallbackIndex < row.length) {
      return row[fallbackIndex];
    }
    return undefined;
  };
}

function matchesFilter_(value, filter) {
  if (filter === undefined || filter === null) return true;

  if (Array.isArray(filter)) {
    if (!filter.length) return true;
    return filter.some(candidate => normalizeString_(candidate) === normalizeString_(value));
  }

  const normalizedFilter = normalizeString_(filter);
  if (!normalizedFilter) return true;

  return normalizeString_(value) === normalizedFilter;
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

function writeRowsAlignedToTarget_(spreadsheetId, sheetName, sourceHeader, rows) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`找不到目標分頁：${sheetName}`);

  const lastColumn = sheet.getLastColumn();
  if (lastColumn < 1) {
    throw new Error(`目標分頁缺少標題列：${sheetName}`);
  }

  const targetHeader = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  if (!targetHeader.length) {
    throw new Error(`目標分頁缺少標題列：${sheetName}`);
  }

  const targetHeaderMap = buildHeaderMap_(targetHeader);
  const columnMappings = sourceHeader.reduce((acc, name, index) => {
    const key = normalizeString_(name);
    if (!key) return acc;
    if (!Object.prototype.hasOwnProperty.call(targetHeaderMap, key)) return acc;
    acc.push({
      sourceIndex: index,
      targetColumn: targetHeaderMap[key] + 1
    });
    return acc;
  }, []);

  if (!columnMappings.length) return;

  const requiredRows = rows.length + 1;
  ensureCapacity_(sheet, requiredRows, sheet.getMaxColumns());

  const maxRows = sheet.getMaxRows();
  columnMappings.forEach(({ targetColumn }) => {
    if (maxRows > 1) {
      sheet.getRange(2, targetColumn, maxRows - 1, 1).clearContent();
    }
  });

  if (!rows.length) return;

  columnMappings.forEach(({ sourceIndex, targetColumn }) => {
    const columnValues = rows.map(row => {
      const value = row[sourceIndex];
      return [value === undefined ? '' : value];
    });
    sheet.getRange(2, targetColumn, columnValues.length, 1).setValues(columnValues);
  });
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

function buildHeaderMap_(header) {
  const map = Object.create(null);
  if (!Array.isArray(header)) return map;

  header.forEach((name, index) => {
    const key = normalizeString_(name);
    if (!key) return;
    if (!Object.prototype.hasOwnProperty.call(map, key)) {
      map[key] = index;
    }
  });

  return map;
}

function normalizeRowLength_(row, targetLength) {
  const normalized = row.slice(0, targetLength);
  while (normalized.length < targetLength) {
    normalized.push('');
  }
  return normalized;
}
