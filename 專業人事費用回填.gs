/**
 * 回填人事費用至各專業中心預算表。
 */

const PERSONNEL_COST_BACKFILL_SOURCE_SPREADSHEET_ID = '1YnuIpuBkZ7lVcvWlIqyMg6vN8-xjgsh-5lRQ2Sv1I3s';

const PERSONNEL_COST_BACKFILL_SOURCE_SHEETS = {
  ICHEF: 'iCHEF 人事費用輸出表',
  QLR: 'QLR 人事費用輸出表',
};

const PERSONNEL_COST_BACKFILL_SHEET_PAIRS = [
  {
    referenceSheetName: '1.1_各事業 TA 預算規劃_人力時間配置',
    targetSheetName: '1.2_各事業 TA 預算規劃_$',
  },
  {
    referenceSheetName: '2.1_事業維運費用_人力時間配置',
    targetSheetName: '2.2_事業維運費用_$',
  },
  {
    referenceSheetName: '3.1_集團共用資源_人力時間配置',
    targetSheetName: '3.2_集團共用資源_$',
  },
];

const PERSONNEL_COST_BACKFILL_TYPE = '人事費用';
const PERSONNEL_COST_BACKFILL_ITEM = '加總';
const PERSONNEL_COST_BACKFILL_KEY_INDEXES = [0, 1, 2, 3, 5];
const PERSONNEL_COST_BACKFILL_GROUP_BY_SUBSIDIARY = {
  TW: 'ICHEF',
  QLR: 'QLR',
};

function addPersonnelCostBackfillMenu_(ui) {
  (ui || SpreadsheetApp.getUi())
    .createMenu('人事費用回填')
    .addItem('回填全部中心', 'backfillPersonnelCostsAll')
    .addSeparator()
    .addItem('僅回填 iCHEF', 'backfillPersonnelCostsICHEF')
    .addItem('僅回填 QLR', 'backfillPersonnelCostsQLR')
    .addToUi();
}

function backfillPersonnelCostsAll() {
  runPersonnelCostBackfill_();
}

function backfillPersonnelCostsICHEF() {
  runPersonnelCostBackfill_('ICHEF');
}

function backfillPersonnelCostsQLR() {
  runPersonnelCostBackfill_('QLR');
}

function runPersonnelCostBackfill_(groupFilter) {
  const spreadsheet = SpreadsheetApp.getActive();
  if (!spreadsheet) {
    throw new Error('找不到目前活頁簿，請在目標試算表中執行人事費用回填。');
  }
  const sourceDataByGroup = loadPersonnelSourceAggregates_(groupFilter);
  PERSONNEL_COST_BACKFILL_SHEET_PAIRS.forEach(function (pair) {
    try {
      backfillPersonnelCostForSheet_(spreadsheet, pair, sourceDataByGroup, groupFilter);
    } catch (error) {
      Logger.log('[PersonnelCostBackfill] ' + spreadsheet.getName() + ' / ' + pair.targetSheetName + ' 發生錯誤：' + error.message);
    }
  });
}

function loadPersonnelSourceAggregates_(groupFilter) {
  const groupsToLoad = groupFilter ? [groupFilter] : Object.keys(PERSONNEL_COST_BACKFILL_SOURCE_SHEETS);
  const spreadsheet = SpreadsheetApp.openById(PERSONNEL_COST_BACKFILL_SOURCE_SPREADSHEET_ID);
  const result = {};

  groupsToLoad.forEach(function (groupKey) {
    const sheetName = PERSONNEL_COST_BACKFILL_SOURCE_SHEETS[groupKey];
    if (!sheetName) {
      return;
    }
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log('[PersonnelCostBackfill] 找不到來源分頁：' + sheetName);
      return;
    }
    result[groupKey] = buildPersonnelSourceAggregate_(sheet);
  });

  return result;
}

function buildPersonnelSourceAggregate_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (!values.length) {
    throw new Error('來源分頁 ' + sheet.getName() + ' 沒有資料列');
  }

  const header = values[0];
  const monthColumns = detectPersonnelMonthColumns_(header);
  if (!monthColumns.length) {
    throw new Error('來源分頁 ' + sheet.getName() + ' 找不到任何月份欄位');
  }

  const aggregated = new Map();
  const rows = values.slice(1);

  rows.forEach(function (row) {
    const info = buildPersonnelKeyFromRow_(row);
    if (!info) return;

    if (!aggregated.has(info.key)) {
      const totalsTemplate = {};
      monthColumns.forEach(function (month) {
        totalsTemplate[month.key] = 0;
      });
      aggregated.set(info.key, {
        codes: info.codes,
        totals: totalsTemplate,
      });
    }

    const bucket = aggregated.get(info.key);
    monthColumns.forEach(function (month) {
      const amount = toNumberStrict_(row[month.index]);
      if (!isNaN(amount)) {
        bucket.totals[month.key] += amount;
      }
    });
  });

  return {
    header: header,
    monthColumns: monthColumns,
    aggregated: aggregated,
  };
}

function backfillPersonnelCostForSheet_(spreadsheet, pair, sourceDataByGroup, groupFilter) {
  const referenceSheet = spreadsheet.getSheetByName(pair.referenceSheetName);
  if (!referenceSheet) {
    Logger.log('[PersonnelCostBackfill] ' + spreadsheet.getName() + ' 找不到參考分頁：' + pair.referenceSheetName);
    return;
  }

  const targetSheet = spreadsheet.getSheetByName(pair.targetSheetName);
  if (!targetSheet) {
    Logger.log('[PersonnelCostBackfill] ' + spreadsheet.getName() + ' 找不到輸出分頁：' + pair.targetSheetName);
    return;
  }

  const referenceKeys = collectPersonnelReferenceKeys_(referenceSheet);
  if (!referenceKeys.size) {
    Logger.log('[PersonnelCostBackfill] ' + spreadsheet.getName() + ' / ' + pair.referenceSheetName + ' 沒有可比對的編碼');
    return;
  }

  const derivedGroups = determineRequiredGroupKeys_(referenceKeys);
  let groupKeysToApply;
  if (groupFilter) {
    groupKeysToApply = [groupFilter];
  } else if (derivedGroups.length > 0) {
    groupKeysToApply = derivedGroups;
  } else {
    groupKeysToApply = Object.keys(sourceDataByGroup);
  }

  groupKeysToApply.forEach(function (groupKey) {
    const sourceData = sourceDataByGroup[groupKey];
    if (!sourceData) {
      Logger.log('[PersonnelCostBackfill] 找不到來源資料：' + groupKey + '，跳過 ' + spreadsheet.getName() + ' / ' + pair.targetSheetName);
      return;
    }
    if (!sourceData.aggregated || sourceData.aggregated.size === 0) {
      Logger.log('[PersonnelCostBackfill] 來源資料為空：' + groupKey + '，跳過 ' + spreadsheet.getName() + ' / ' + pair.targetSheetName);
      return;
    }
    applyPersonnelAggregatesToTarget_(
      targetSheet,
      referenceKeys,
      sourceData,
      spreadsheet.getName() + ' / ' + pair.targetSheetName + ' [' + groupKey + ']'
    );
  });
}

function collectPersonnelReferenceKeys_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return new Map();
  }
  const columnCount = Math.max(PERSONNEL_COST_BACKFILL_KEY_INDEXES[PERSONNEL_COST_BACKFILL_KEY_INDEXES.length - 1] + 1, sheet.getLastColumn());
  const values = sheet.getRange(2, 1, lastRow - 1, columnCount).getValues();
  const map = new Map();

  values.forEach(function (row) {
    const info = buildPersonnelKeyFromRow_(row);
    if (!info) return;
    if (!map.has(info.key)) {
      map.set(info.key, info.codes);
    }
  });

  return map;
}

function determineRequiredGroupKeys_(referenceKeyMap) {
  const groups = new Set();
  referenceKeyMap.forEach(function (codes) {
    if (!codes || !codes.length) {
      return;
    }
    const subsidiary = codes[0];
    const groupKey = PERSONNEL_COST_BACKFILL_GROUP_BY_SUBSIDIARY[subsidiary];
    if (groupKey) {
      groups.add(groupKey);
    }
  });
  return Array.from(groups);
}

function applyPersonnelAggregatesToTarget_(sheet, referenceKeys, sourceData, contextLabel) {
  const lastColumn = sheet.getLastColumn();
  if (!lastColumn) {
    throw new Error('目標分頁缺少表頭：' + contextLabel);
  }

  const header = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const monthColumns = detectPersonnelMonthColumns_(header);
  if (!monthColumns.length) {
    throw new Error('目標分頁缺少月份欄位：' + contextLabel);
  }

  const costTypeIndex = header.indexOf('費用類型');
  const costItemIndex = header.indexOf('費用項目');
  const yearTotalIndex = header.indexOf('2026 year');

  const rowCount = sheet.getLastRow() - 1;
  const dataValues = rowCount > 0 ? sheet.getRange(2, 1, rowCount, lastColumn).getValues() : [];
  const existingRowMap = new Map();
  const blankRowQueue = [];
  let lastDataRow = 1;

  dataValues.forEach(function (row, idx) {
    const rowNumber = idx + 2;
    const info = buildPersonnelKeyFromRow_(row);
    const costItemValue = costItemIndex >= 0 ? normalizeMaybe_(row[costItemIndex]) : '';
    const hasMonthValue = monthColumns.some(function (month) {
      return hasMeaningfulValue_(row[month.index]);
    });
    const hasCodes = info ? info.codes.some(function (code) {
      return !!code;
    }) : false;
    const costTypeValue = costTypeIndex >= 0 ? normalizeMaybe_(row[costTypeIndex]) : '';
    const hasAnyValue = hasMonthValue || hasCodes || !!costItemValue || !!costTypeValue;

    if (info && costItemValue === PERSONNEL_COST_BACKFILL_ITEM) {
      existingRowMap.set(info.key, rowNumber);
      if (hasAnyValue) {
        lastDataRow = rowNumber;
      }
      return;
    }

    if (!hasMonthValue && !hasCodes && !costItemValue && !costTypeValue) {
      blankRowQueue.push(rowNumber);
      return;
    }

    if (hasAnyValue) {
      lastDataRow = rowNumber;
    }
  });

  let availableBlankRows = blankRowQueue.filter(function (rowNumber) {
    return rowNumber > lastDataRow;
  });

  const updates = [];
  let updatedCount = 0;

  referenceKeys.forEach(function (codes, key) {
    const aggregate = sourceData.aggregated.get(key);
    if (!aggregate) {
      return;
    }

    const totals = monthColumns.map(function (month) {
      const value = Object.prototype.hasOwnProperty.call(aggregate.totals, month.key)
        ? aggregate.totals[month.key]
        : 0;
      return sanitizeTotal_(value);
    });

    const hasAmount = totals.some(function (value) {
      return value !== 0;
    });

    let rowNumber = existingRowMap.get(key);
    if (!rowNumber && !hasAmount) {
      return;
    }

    if (!rowNumber && availableBlankRows.length) {
      rowNumber = availableBlankRows.shift();
    }
    if (!rowNumber) {
      rowNumber = ensureAppendRow_(sheet);
    }

    const rowValues = getRowSnapshot_(dataValues, rowNumber, lastColumn);
    PERSONNEL_COST_BACKFILL_KEY_INDEXES.forEach(function (columnIndex, position) {
      if (columnIndex < rowValues.length) {
        rowValues[columnIndex] = codes[position] || '';
      }
    });
    if (costTypeIndex >= 0) {
      rowValues[costTypeIndex] = PERSONNEL_COST_BACKFILL_TYPE;
    }
    if (costItemIndex >= 0) {
      rowValues[costItemIndex] = PERSONNEL_COST_BACKFILL_ITEM;
    }
    monthColumns.forEach(function (month, idx) {
      rowValues[month.index] = totals[idx];
    });
    if (yearTotalIndex >= 0) {
      const annualTotal = monthColumns.reduce(function (sum, month, idx) {
        if (typeof month.key === 'string' && month.key.indexOf('2026-') === 0) {
          return sum + (Number(totals[idx]) || 0);
        }
        return sum;
      }, 0);
      rowValues[yearTotalIndex] = sanitizeTotal_(annualTotal);
    }

    updates.push({ rowNumber: rowNumber, values: rowValues });
    updateCachedRow_(dataValues, rowNumber, rowValues);
    updatedCount += 1;
  });

  if (!updates.length) {
    Logger.log('[PersonnelCostBackfill] ' + contextLabel + ' 無需更新。');
    return;
  }

  writeRowUpdates_(sheet, updates, lastColumn);
  Logger.log('[PersonnelCostBackfill] ' + contextLabel + ' 完成回填 ' + updatedCount + ' 筆。');
}

function buildPersonnelKeyFromRow_(row) {
  const codes = PERSONNEL_COST_BACKFILL_KEY_INDEXES.map(function (columnIndex) {
    const value = columnIndex < row.length ? row[columnIndex] : '';
    return normalizeMaybe_(value);
  });
  const hasValue = codes.some(function (code) {
    return !!code;
  });

  if (!hasValue) {
    return null;
  }

  return {
    key: codes.join('||'),
    codes: codes,
  };
}

function detectPersonnelMonthColumns_(header) {
  const result = [];
  header.forEach(function (value, index) {
    const monthKey = computeMonthKey_(value);
    if (monthKey) {
      result.push({
        index: index,
        key: monthKey,
        header: value,
      });
    }
  });
  return result;
}

function computeMonthKey_(value) {
  if (value instanceof Date) {
    return value.getFullYear() + '-' + padTwo_(value.getMonth() + 1);
  }

  const normalized = normalizeMaybe_(value).replace(/\s+/g, '');
  if (!normalized) {
    return null;
  }

  const match = normalized.match(/^(\d{4})[\/\-\.年]?(\d{1,2})月?$/);
  if (!match) {
    return null;
  }

  const year = match[1];
  const month = padTwo_(Number(match[2]));
  return year + '-' + month;
}

function padTwo_(value) {
  const num = Number(value);
  if (!isFinite(num)) {
    return String(value);
  }
  return num < 10 ? '0' + num : String(num);
}

function getRowSnapshot_(dataValues, rowNumber, columnCount) {
  const index = rowNumber - 2;
  if (index >= 0 && index < dataValues.length) {
    const clone = dataValues[index].slice();
    while (clone.length < columnCount) {
      clone.push('');
    }
    return clone;
  }
  return new Array(columnCount).fill('');
}

function updateCachedRow_(dataValues, rowNumber, rowValues) {
  const index = rowNumber - 2;
  if (index >= 0 && index < dataValues.length) {
    dataValues[index] = rowValues.slice();
  }
}

function writeRowUpdates_(sheet, updates, columnCount) {
  updates.sort(function (a, b) {
    return a.rowNumber - b.rowNumber;
  });

  let start = 0;
  while (start < updates.length) {
    let end = start + 1;
    const baseRow = updates[start].rowNumber;
    while (end < updates.length && updates[end].rowNumber === baseRow + (end - start)) {
      end += 1;
    }
    const rows = updates.slice(start, end).map(function (entry) {
      const values = entry.values.slice();
      while (values.length < columnCount) {
        values.push('');
      }
      return values;
    });
    sheet.getRange(updates[start].rowNumber, 1, rows.length, columnCount).setValues(rows);
    start = end;
  }
}

function ensureAppendRow_(sheet) {
  const targetRow = sheet.getLastRow() + 1;
  const maxRows = sheet.getMaxRows();
  if (targetRow > maxRows) {
    sheet.insertRowsAfter(maxRows, targetRow - maxRows);
  }
  return targetRow;
}

function hasMeaningfulValue_(value) {
  if (value === null || value === undefined) {
    return false;
  }
  if (value instanceof Date) {
    return true;
  }
  if (typeof value === 'number') {
    return value !== 0;
  }
  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (!trimmed) {
      return false;
    }
    return trimmed !== '0' && trimmed !== '0.0' && trimmed !== '0.00';
  }
  return true;
}

function toNumberStrict_(value) {
  if (value === null || value === undefined) {
    return 0;
  }
  if (typeof value === 'number') {
    return isFinite(value) ? value : NaN;
  }
  if (typeof value === 'string') {
    const normalized = value.replace(/[,\\s]/g, '');
    if (!normalized) {
      return 0;
    }
    const parsed = Number(normalized);
    return isFinite(parsed) ? parsed : NaN;
  }
  return NaN;
}

function sanitizeTotal_(value) {
  if (typeof value !== 'number') {
    const parsed = Number(value);
    if (!isFinite(parsed)) {
      return 0;
    }
    value = parsed;
  }
  if (!isFinite(value)) {
    return 0;
  }
  const rounded = Math.round(value * 100) / 100;
  return Math.abs(rounded) === 0 ? 0 : rounded;
}

function normalizeMaybe_(value) {
  if (typeof normalizeString_ === 'function') {
    return normalizeString_(value);
  }
  if (value === null || value === undefined) {
    return '';
  }
  if (typeof value === 'string') {
    return value.trim();
  }
  if (typeof value === 'number' && isFinite(value)) {
    return String(value);
  }
  return String(value).trim();
}
