/**
 * 依據指定條件整理「2.2_事業維運費用_$」的資料，
 * 將每月合計拆分為 9.25% / 90.75%，並寫入「維運攤分」分頁。
 */
function addMaintenanceMenu_(ui) {
  (ui || SpreadsheetApp.getUi())
    .createMenu('維運攤分')
    .addItem('執行攤分', 'distributeMaintenanceCosts')
    .addToUi();
}

function distributeMaintenanceCosts() {
  const SPREADSHEET_ID = '1mYeC6DpUqFvnce9leb-098wSYCGqfOR5_HGDM2bcbEc';
  const SOURCE_SHEET_NAME = '2.2_事業維運費用_$';
  const TARGET_SHEET_NAME = '維運攤分';

  const FILTER_CRITERIA = {
    '子公司': 'TW',
    '事業單位': 'Corporation',
    'Si 編號': 'Maintenance',
    'TA 編號': 'Baseline',
    '價值鏈專案預算代號': 'NA',
    '費用類型': '非人事費用',
  };

  const SPLIT_CONFIG = [
    {
      share: 0.0925,
      fields: {
        '子公司': 'TW',
        '事業單位': 'OMO',
        'Si 編號': 'Maintenance',
        'TA 編號': 'Baseline',
        '價值鏈專案預算代號': 'NA',
        '費用單位': '執行長室 : 客戶價值中心 : 客戶價值中心',
        '費用類型': '非人事費用',
        '費用項目': '維運費用分攤',
      },
    },
    {
      share: 0.9075,
      fields: {
        '子公司': 'TW',
        '事業單位': 'POS',
        'Si 編號': 'Maintenance',
        'TA 編號': 'Baseline',
        '價值鏈專案預算代號': 'NA',
        '費用單位': '執行長室 : 客戶價值中心 : 客戶價值中心',
        '費用類型': '非人事費用',
        '費用項目': '維運費用分攤',
      },
    },
  ];

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sourceSheet = ss.getSheetByName(SOURCE_SHEET_NAME);
  if (!sourceSheet) {
    throw new Error('找不到來源分頁：' + SOURCE_SHEET_NAME);
  }
  const targetSheet = ss.getSheetByName(TARGET_SHEET_NAME);
  if (!targetSheet) {
    throw new Error('找不到目的分頁：' + TARGET_SHEET_NAME);
  }

  const sourceValues = sourceSheet.getDataRange().getValues();
  if (sourceValues.length < 2) {
    Logger.log('來源資料不存在資料列。');
    removeExistingAllocations_(targetSheet, null);
    return;
  }

  const sourceHeaders = sourceValues[0];
  const sourceHeaderMap = buildHeaderMap_(sourceHeaders);
  const filteredRows = sourceValues.slice(1).filter(function (row) {
    return matchesCriteria_(row, sourceHeaderMap, FILTER_CRITERIA);
  });

  if (filteredRows.length === 0) {
    Logger.log('沒有符合條件的來源資料。');
    removeExistingAllocations_(targetSheet, SPLIT_CONFIG[0].fields);
    return;
  }

  const reservedHeaderNames = collectReservedHeaders_(FILTER_CRITERIA, SPLIT_CONFIG);
  const monthColumns = detectMonthColumns_(sourceHeaders, filteredRows, reservedHeaderNames);
  if (monthColumns.length === 0) {
    throw new Error('找不到任何看起來像月份的欄位，請確認表頭設定。');
  }

  const monthlyTotals = monthColumns.map(function (col) {
    return {
      header: col.header,
      total: filteredRows.reduce(function (sum, row) {
        return sum + toNumber_(row[col.index]);
      }, 0),
    };
  });

  const hasNonZero = monthlyTotals.some(function (item) {
    return Math.abs(item.total) > 0;
  });
  if (!hasNonZero) {
    Logger.log('符合條件的資料所有月份皆為 0，僅清除既有資料。');
    removeExistingAllocations_(targetSheet, SPLIT_CONFIG[0].fields);
    return;
  }

  const targetValues = targetSheet.getDataRange().getValues();
  if (targetValues.length === 0) {
    throw new Error('目的分頁缺少表頭列。');
  }
  const targetHeaders = targetValues[0];
  const targetHeaderMap = buildHeaderMap_(targetHeaders);

  removeExistingAllocations_(targetSheet, null, targetHeaderMap);

  const shares = SPLIT_CONFIG.map(function (config) {
    return config.share;
  });
  const monthlyAllocations = monthlyTotals.map(function (monthInfo) {
    return allocateByShares_(monthInfo.total, shares);
  });

  const outputRows = SPLIT_CONFIG.map(function (config, configIndex) {
    const row = new Array(targetHeaders.length).fill('');
    Object.keys(config.fields).forEach(function (key) {
      if (key in targetHeaderMap) {
        row[targetHeaderMap[key]] = config.fields[key];
      }
    });
    monthlyTotals.forEach(function (monthInfo, monthIndex) {
      if (!(monthInfo.header in targetHeaderMap)) {
        Logger.log('目的分頁缺少欄位：' + monthInfo.header + '，略過該欄位。');
        return;
      }
      const idx = targetHeaderMap[monthInfo.header];
      row[idx] = monthlyAllocations[monthIndex][configIndex];
    });
    return row;
  });

  const targetStartRow = targetSheet.getLastRow() + 1;
  targetSheet
    .getRange(targetStartRow, 1, outputRows.length, targetHeaders.length)
    .setValues(outputRows);
}

function buildHeaderMap_(headers) {
  var map = {};
  headers.forEach(function (header, index) {
    if (header && !(header in map)) {
      map[header] = index;
    }
  });
  return map;
}

function matchesCriteria_(row, headerMap, criteria) {
  return Object.keys(criteria).every(function (key) {
    if (!(key in headerMap)) {
      throw new Error('來源表頭缺少欄位：' + key);
    }
    const value = normalizeString_(row[headerMap[key]]);
    return value === normalizeString_(criteria[key]);
  });
}

function detectMonthColumns_(headers, rows, reservedHeaders) {
  const reserved = reservedHeaders || new Set();
  return headers.reduce(function (acc, header, index) {
    if (!header || reserved.has(header)) {
      return acc;
    }
    if (!isMonthlyHeader_(header)) {
      return acc;
    }
    const hasNumeric = rows.some(function (row) {
      return typeof row[index] === 'number' && !isNaN(row[index]);
    });
    if (hasNumeric) {
      acc.push({ header: header, index: index });
    }
    return acc;
  }, []);
}

function isMonthlyHeader_(header) {
  const MONTH_PATTERN =
    /(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec|Q[1-4]|月|20\d{2}|總計|合計|總額|Total|Sum)/i;
  return MONTH_PATTERN.test(header);
}

function toNumber_(value) {
  if (typeof value === 'number') {
    return value;
  }
  if (typeof value === 'string' && value.trim() !== '') {
    const parsed = parseFloat(value.replace(/,/g, ''));
    return isNaN(parsed) ? 0 : parsed;
  }
  return 0;
}

function roundAmount_(value) {
  return Math.round(value * 100) / 100;
}

function allocateByShares_(total, shares) {
  var allocations = [];
  var remaining = total;
  for (var i = 0; i < shares.length; i++) {
    var amount;
    if (i === shares.length - 1) {
      amount = roundAmount_(remaining);
    } else {
      amount = roundAmount_(total * shares[i]);
      remaining = roundAmount_(remaining - amount);
    }
    allocations.push(amount);
  }
  return allocations;
}

function collectReservedHeaders_(criteria, splitConfig) {
  const reserved = new Set(Object.keys(criteria));
  splitConfig.forEach(function (config) {
    Object.keys(config.fields).forEach(function (key) {
      reserved.add(key);
    });
  });
  reserved.add('費用項目');
  return reserved;
}

function removeExistingAllocations_(sheet, referenceFields, headerMap) {
  const dataRange = sheet.getDataRange();
  if (dataRange.getNumRows() < 2) {
    return;
  }
  const values = dataRange.getValues();
  const headers = values[0];
  const map = headerMap || buildHeaderMap_(headers);

  const matchFields = referenceFields || {
    '子公司': 'TW',
    'Si 編號': 'Maintenance',
    'TA 編號': 'Baseline',
    '價值鏈專案預算代號': 'NA',
    '費用項目': '維運費用分攤',
  };

  const rowsToDelete = [];
  for (var i = 1; i < values.length; i++) {
    if (rowMatches_(values[i], map, matchFields)) {
      rowsToDelete.push(i + 1); // 1-based row index
    }
  }
  rowsToDelete
    .sort(function (a, b) {
      return b - a;
    })
    .forEach(function (rowIndex) {
      sheet.deleteRow(rowIndex);
    });
}

function rowMatches_(row, headerMap, fields) {
  return Object.keys(fields).every(function (key) {
    if (!(key in headerMap)) {
      return false;
    }
    return normalizeString_(row[headerMap[key]]) === normalizeString_(fields[key]);
  });
}

function normalizeString_(value) {
  if (value === null || value === undefined) {
    return '';
  }
  if (typeof value === 'string') {
    return value.trim();
  }
  if (typeof value === 'number') {
    return String(value);
  }
  return String(value).trim();
}
