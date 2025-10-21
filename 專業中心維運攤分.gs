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
  // TODO: 使用前請更新對應的 SPREADSHEET_ID。
  const SPREADSHEET_ID = '1mYeC6DpUqFvnce9leb-098wSYCGqfOR5_HGDM2bcbEc';
  const SOURCE_SHEET_NAME = '2.2_事業維運費用_$';
  const TARGET_SHEET_NAME = '維運攤分';

  const BASE_FILTER = {
    '子公司': 'TW',
    '事業單位': 'Corporation',
    'Si 編號': 'Maintenance',
    'TA 編號': 'Baseline',
    '價值鏈專案預算代號': 'NA',
  };

  const ALLOWED_COST_TYPES = ['非人事費用', '人事費用'];

  const SPLIT_CONFIG_TEMPLATE = [
    {
      share: 0.0925,
      fields: {
        '子公司': 'TW',
        '事業單位': 'OMO',
        'Si 編號': 'Maintenance',
        'TA 編號': 'Baseline',
        '價值鏈專案預算代號': 'NA',
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
        '費用類型': '非人事費用',
        '費用項目': '維運費用分攤',
      },
    },
  ];

  const COST_UNIT_MAPPING = {
    '執行長室 : 管理中心': '執行長室 : 管理中心 : 管理中心',
    '執行長室 : 財務中心': '執行長室 : 財務中心 : 財務中心',
    '執行長室 : 研發中心': '執行長室 : 研發中心 : 研發中心',
    '執行長室 : 行銷營運中心': '執行長室 : 行銷營運中心 : 行銷營運中心',
    '執行長室 : 客戶價值中心': '執行長室 : 客戶價值中心 : 客戶價值中心',
    '執行長室 : 策略資料中心': '執行長室 : 策略資料中心 : 策略資料中心',
  };

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
    removeExistingAllocations_(targetSheet, {
      '子公司': 'TW',
      'Si 編號': 'Maintenance',
      'TA 編號': 'Baseline',
      '價值鏈專案預算代號': 'NA',
      '費用項目': '維運費用分攤',
    });
    return;
  }

  const sourceHeaders = sourceValues[0];
  const sourceHeaderMap = buildHeaderMap_(sourceHeaders);
  assertHeadersPresent_(sourceHeaderMap, ['費用類型', '費用單位'], '來源分頁');
  const baseRows = sourceValues.slice(1).filter(function (row) {
    return matchesCriteria_(row, sourceHeaderMap, BASE_FILTER);
  });

  if (baseRows.length === 0) {
    Logger.log('沒有符合條件的來源資料 (Corporation)。');
  }

  const fallbackRows =
    baseRows.length > 0
      ? baseRows
      : sourceValues.slice(1).filter(function (row) {
          const subsidiary = normalizeString_(row[sourceHeaderMap['子公司']]);
          const businessUnit = normalizeString_(row[sourceHeaderMap['事業單位']]);
          return subsidiary === 'TW' && (businessUnit === 'POS' || businessUnit === 'OMO');
        });

  if (fallbackRows.length === 0) {
    Logger.log('沒有可供分攤或直接加總的維運資料。');
    removeExistingAllocations_(targetSheet, {
      '子公司': 'TW',
      'Si 編號': 'Maintenance',
      'TA 編號': 'Baseline',
      '價值鏈專案預算代號': 'NA',
      '費用項目': '維運費用分攤',
    });
    return;
  }

  const targetCostUnit = determineTargetCostUnit_(baseRows, fallbackRows, sourceHeaderMap, COST_UNIT_MAPPING);
  const splitConfig = applyCostUnitToSplitConfig_(SPLIT_CONFIG_TEMPLATE, targetCostUnit);

  const reservedHeaderNames = collectReservedHeaders_(BASE_FILTER, splitConfig);
  const monthSourceRows = baseRows.length > 0 ? baseRows : fallbackRows;
  const monthColumns = detectMonthColumns_(sourceHeaders, monthSourceRows, reservedHeaderNames);
  if (monthColumns.length === 0) {
    throw new Error('找不到任何看起來像月份的欄位，請確認表頭設定。');
  }

  const costTypeTotals = {};
  ALLOWED_COST_TYPES.forEach(function (costType) {
    costTypeTotals[costType] = monthColumns.map(function () {
      return 0;
    });
  });
  baseRows.forEach(function (row) {
    const costType = normalizeString_(row[sourceHeaderMap['費用類型']]);
    if (!ALLOWED_COST_TYPES.includes(costType)) {
      return;
    }
    monthColumns.forEach(function (monthCol, index) {
      costTypeTotals[costType][index] += toNumber_(row[monthCol.index]);
    });
  });

  const targetValues = targetSheet.getDataRange().getValues();
  if (targetValues.length === 0) {
    throw new Error('目的分頁缺少表頭列。');
  }
  const targetHeaders = targetValues[0];
  const targetHeaderMap = buildHeaderMap_(targetHeaders);

  removeExistingAllocations_(targetSheet, {
    '子公司': 'TW',
    'Si 編號': 'Maintenance',
    'TA 編號': 'Baseline',
    '價值鏈專案預算代號': 'NA',
    '費用項目': '維運費用分攤',
  }, targetHeaderMap);

  const directPersonnelTotals = collectDirectMaintenancePersonnelTotals_(
    sourceValues.slice(1),
    sourceHeaderMap,
    monthColumns,
    splitConfig
  );
  const directNonPersonnelTotals = collectDirectMaintenanceNonPersonnelTotals_(
    sourceValues.slice(1),
    sourceHeaderMap,
    monthColumns,
    splitConfig
  );

  const shares = splitConfig.map(function (config) {
    return config.share;
  });

  const outputRows = [];
  ALLOWED_COST_TYPES.forEach(function (costType) {
    const totals = costTypeTotals[costType];
    const hasTotals = hasNonZero_(totals);
    const allocations = hasTotals
      ? totals.map(function (total) {
          return allocateByShares_(total, shares);
        })
      : monthColumns.map(function () {
          return shares.map(function () {
            return 0;
          });
        });

    splitConfig.forEach(function (config, configIndex) {
      const row = new Array(targetHeaders.length).fill('');
      Object.keys(config.fields).forEach(function (key) {
        if (key in targetHeaderMap) {
          row[targetHeaderMap[key]] = config.fields[key];
        }
      });
      if ('費用類型' in targetHeaderMap) {
        row[targetHeaderMap['費用類型']] = costType;
      }
      let hasRowData = false;
      monthColumns.forEach(function (monthCol, monthIndex) {
        if (!(monthCol.header in targetHeaderMap)) {
          Logger.log('目的分頁缺少欄位：' + monthCol.header + '，略過該欄位。');
          return;
        }
        let value = allocations[monthIndex][configIndex] || 0;
        if (costType === '人事費用') {
          const directKey = config.fields['子公司'] + '|' + config.fields['事業單位'];
          const directTotals = directPersonnelTotals.get(directKey);
          if (directTotals) {
            value = roundAmount_(value + directTotals[monthIndex]);
          }
        } else if (costType === '非人事費用') {
          const directKey = config.fields['子公司'] + '|' + config.fields['事業單位'];
          const directTotals = directNonPersonnelTotals.get(directKey);
          if (directTotals) {
            value = roundAmount_(value + directTotals[monthIndex]);
          }
        }
        if (Math.abs(value) > 0) {
          hasRowData = true;
        }
        row[targetHeaderMap[monthCol.header]] = value;
      });
      if (hasRowData) {
        outputRows.push(row);
      }
    });
  });

  if (outputRows.length === 0) {
    Logger.log('沒有需要輸出的維運攤分資料，已清除既有資料列。');
    return;
  }

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

function applyCostUnitToSplitConfig_(template, costUnit) {
  return template.map(function (config) {
    return {
      share: config.share,
      fields: Object.assign({}, config.fields, {
        '費用單位': costUnit,
      }),
    };
  });
}

function determineTargetCostUnit_(primaryRows, fallbackRows, headerMap, costUnitMap) {
  var mapped = tryResolveCostUnitFromRows_(primaryRows, headerMap, costUnitMap);
  if (mapped) {
    return mapped;
  }
  mapped = tryResolveCostUnitFromRows_(fallbackRows, headerMap, costUnitMap);
  if (mapped) {
    return mapped;
  }
  throw new Error('找不到可辨識的維運費用單位分類，請確認來源資料。');
}

function tryResolveCostUnitFromRows_(rows, headerMap, costUnitMap) {
  if (!rows || rows.length === 0) {
    return '';
  }
  const index = headerMap['費用單位'];
  const prefixes = new Set();
  rows.forEach(function (row) {
    const prefix = extractCostUnitPrefix_(row[index]);
    if (prefix) {
      prefixes.add(prefix);
    }
  });
  if (prefixes.size === 0) {
    return '';
  }
  if (prefixes.size > 1) {
    throw new Error('來源資料包含多個維運費用單位分類，請確認：' + Array.from(prefixes).join(', '));
  }
  const prefix = prefixes.values().next().value;
  const mapped = costUnitMap[prefix];
  if (!mapped) {
    throw new Error('無法辨識的維運費用單位分類：' + prefix);
  }
  return mapped;
}

function extractCostUnitPrefix_(value) {
  const normalized = normalizeString_(value);
  if (!normalized) {
    return '';
  }
  const segments = normalized
    .split(/[:：]/)
    .map(function (part) {
      return part.trim();
    })
    .filter(function (part) {
      return part !== '';
    });
  if (segments.length < 2) {
    return '';
  }
  return segments[0] + ' : ' + segments[1];
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

function collectDirectMaintenancePersonnelTotals_(rows, headerMap, monthColumns, splitConfig) {
  var allowedCombos = new Set(
    splitConfig.map(function (config) {
      return config.fields['子公司'] + '|' + config.fields['事業單位'];
    })
  );
  var totalsMap = new Map();
  rows.forEach(function (row) {
    var costType = normalizeString_(row[headerMap['費用類型']]);
    if (costType !== '人事費用') {
      return;
    }
    var subsidiary = normalizeString_(row[headerMap['子公司']]);
    var businessUnit = normalizeString_(row[headerMap['事業單位']]);
    if (!subsidiary || !businessUnit) {
      return;
    }
    var comboKey = subsidiary + '|' + businessUnit;
    if (!allowedCombos.has(comboKey)) {
      return;
    }
    accumulateMonthlyTotals_(totalsMap, comboKey, monthColumns, row);
  });
  return totalsMap;
}

function collectDirectMaintenanceNonPersonnelTotals_(rows, headerMap, monthColumns, splitConfig) {
  var allowedCombos = new Set(
    splitConfig.map(function (config) {
      return config.fields['子公司'] + '|' + config.fields['事業單位'];
    })
  );
  var totalsMap = new Map();
  rows.forEach(function (row) {
    var costType = normalizeString_(row[headerMap['費用類型']]);
    if (costType !== '非人事費用') {
      return;
    }
    var subsidiary = normalizeString_(row[headerMap['子公司']]);
    var businessUnit = normalizeString_(row[headerMap['事業單位']]);
    if (!subsidiary || !businessUnit) {
      return;
    }
    if (businessUnit !== 'POS' && businessUnit !== 'OMO') {
      return;
    }
    var comboKey = subsidiary + '|' + businessUnit;
    if (!allowedCombos.has(comboKey)) {
      return;
    }
    accumulateMonthlyTotals_(totalsMap, comboKey, monthColumns, row);
  });
  return totalsMap;
}

function accumulateMonthlyTotals_(bucketMap, unitKey, monthColumns, row) {
  if (!bucketMap.has(unitKey)) {
    bucketMap.set(
      unitKey,
      monthColumns.map(function () {
        return 0;
      })
    );
  }
  const totals = bucketMap.get(unitKey);
  monthColumns.forEach(function (monthCol, index) {
    totals[index] += toNumber_(row[monthCol.index]);
  });
}

function hasNonZero_(values) {
  return values.some(function (value) {
    return Math.abs(value) > 0;
  });
}

function assertHeadersPresent_(headerMap, headerNames, contextLabel) {
  headerNames.forEach(function (name) {
    if (!(name in headerMap)) {
      throw new Error((contextLabel || '工作表') + '缺少必要欄位：' + name);
    }
  });
}
