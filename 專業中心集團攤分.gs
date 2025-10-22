/**
 * 依據指定條件整理「3.2_集團共用資源_$」的資料，
 * 將每月合計拆分為 8.84% / 86.65% / 4.51%，並寫入「集團攤分」分頁。
 */
function addGroupAllocationMenu_(ui) {
  (ui || SpreadsheetApp.getUi())
    .createMenu('集團攤分')
    .addItem('執行攤分', 'distributeGroupAllocations')
    .addToUi();
}

const GROUP_ALLOCATION_CONFIGS_ = [
  {
    unitSet: new Set([
      '執行長室',
      '執行長室 : 財務中心',
      '執行長室 : 稽核部',
      '執行長室 : 管理中心',
      '執行長室 : 策略資料中心',
    ]),
    splitConfig: [
      { share: 0.0884, fields: { '子公司': 'TW', '事業單位': 'OMO' } },
      { share: 0.8665, fields: { '子公司': 'TW', '事業單位': 'POS' } },
      { share: 0.0451, fields: { '子公司': 'QLR', '事業單位': 'OMO' } },
    ],
  },
  {
    unitSet: new Set([
      '執行長室 : 客戶價值中心',
      '執行長室 : 研發中心',
      '執行長室 : 行銷營運中心',
    ]),
    splitConfig: [
      { share: 0.0925, fields: { '子公司': 'TW', '事業單位': 'OMO' } },
      { share: 0.9075, fields: { '子公司': 'TW', '事業單位': 'POS' } },
    ],
  },
];

const GROUP_ALLOCATION_VALID_UNITS_ = (function () {
  var result = new Set();
  GROUP_ALLOCATION_CONFIGS_.forEach(function (group) {
    group.unitSet.forEach(function (unit) {
      result.add(unit);
    });
  });
  return result;
})();

function distributeGroupAllocations() {
  const SOURCE_SHEET_NAME = '3.2_集團共用資源_$';
  const TARGET_SHEET_NAME = '集團攤分';

  const BASE_FILTER = {
    '子公司': 'TW',
    '事業單位': 'Corporation',
    'Si 編號': 'Corporation',
    'TA 編號': 'Corporation',
  };

  const VALID_PROJECT_CODES = new Set(['NA', '']);
  const VALID_UNITS = GROUP_ALLOCATION_VALID_UNITS_;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
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
      'Si 編號': 'Corporation',
      'TA 編號': 'Corporation',
      '費用項目': '集團費用分攤',
    });
    return;
  }

  const sourceHeaders = sourceValues[0];
  const sourceHeaderMap = buildHeaderMap_(sourceHeaders);
  assertHeadersPresent_(sourceHeaderMap, [
    '子公司',
    '事業單位',
    'Si 編號',
    'TA 編號',
    '價值鏈專案預算代號',
    '費用單位',
    '費用類型',
  ], '來源分頁');

  const reservedHeaders = new Set(Object.keys(BASE_FILTER));
  ['價值鏈專案預算代號', '費用單位', '費用類型', '費用項目'].forEach(function (key) {
    reservedHeaders.add(key);
  });

  const monthColumns = detectMonthColumns_(sourceHeaders, sourceValues.slice(1), reservedHeaders);
  if (monthColumns.length === 0) {
    throw new Error('找不到任何看起來像月份的欄位，請確認表頭設定。');
  }

  const costTypeBuckets = {
    '非人事費用': new Map(),
    '人事費用': new Map(),
  };

  const filteredRows = sourceValues.slice(1).filter(function (row) {
    if (!matchesCriteria_(row, sourceHeaderMap, BASE_FILTER)) {
      return false;
    }
    const projectCode = normalizeString_(row[sourceHeaderMap['價值鏈專案預算代號']]).toUpperCase();
    if (!VALID_PROJECT_CODES.has(projectCode)) {
      return false;
    }
    const rawUnit = normalizeString_(row[sourceHeaderMap['費用單位']]);
    const unitKey = normalizeUnitKey_(rawUnit);
    if (!VALID_UNITS.has(unitKey)) {
      return false;
    }
    const costType = normalizeString_(row[sourceHeaderMap['費用類型']]);
    if (!(costType in costTypeBuckets)) {
      return false;
    }
    accumulateMonthlyTotals_(costTypeBuckets[costType], unitKey, monthColumns, row);
    return true;
  });

  if (filteredRows.length === 0) {
    Logger.log('沒有符合條件的來源資料。');
    removeExistingAllocations_(targetSheet, {
      'Si 編號': 'Corporation',
      'TA 編號': 'Corporation',
      '費用項目': '集團費用分攤',
    });
    return;
  }

  const targetValues = targetSheet.getDataRange().getValues();
  if (targetValues.length === 0) {
    throw new Error('目的分頁缺少表頭列。');
  }
  const targetHeaders = targetValues[0];
  const targetHeaderMap = buildHeaderMap_(targetHeaders);
  assertHeadersPresent_(targetHeaderMap, [
    '子公司',
    '事業單位',
    'Si 編號',
    'TA 編號',
    '價值鏈專案預算代號',
    '費用單位',
    '費用類型',
    '費用項目',
  ], '集團攤分');

  removeExistingAllocations_(targetSheet, {
    'Si 編號': 'Corporation',
    'TA 編號': 'Corporation',
    '費用項目': '集團費用分攤',
  }, targetHeaderMap);

  const directPersonnelTotals = collectDirectPersonnelTotals_(
    sourceValues.slice(1),
    sourceHeaderMap,
    monthColumns
  );

  const outputRows = [];
  Object.keys(costTypeBuckets).forEach(function (costType) {
    const unitMap = costTypeBuckets[costType];
    unitMap.forEach(function (monthlyTotals, unitKey) {
      const splitConfig = getSplitConfigForUnit_(unitKey);
      if (!splitConfig || splitConfig.length === 0) {
        return;
      }
      const hasDirectContribution = splitConfig.some(function (config) {
        const directKey = config.fields['子公司'] + '|' + config.fields['事業單位'] + '|' + unitKey;
        return directPersonnelTotals.has(directKey);
      });
      if (!hasNonZero_(monthlyTotals) && !hasDirectContribution) {
        return;
      }
      const allocations = monthlyTotals.map(function (total) {
        return allocateByShares_(total, splitConfig.map(function (config) {
          return config.share;
        }));
      });

      splitConfig.forEach(function (config, index) {
        const row = new Array(targetHeaders.length).fill('');
        row[targetHeaderMap['子公司']] = config.fields['子公司'];
        row[targetHeaderMap['事業單位']] = config.fields['事業單位'];
        row[targetHeaderMap['Si 編號']] = 'Corporation';
        row[targetHeaderMap['TA 編號']] = 'Corporation';
        if ('價值鏈專案預算代號' in targetHeaderMap) {
          row[targetHeaderMap['價值鏈專案預算代號']] = 'NA';
        }
        if ('費用單位' in targetHeaderMap) {
          row[targetHeaderMap['費用單位']] = getGroupAllocationUnitDisplay_(unitKey);
        }
        if ('費用類型' in targetHeaderMap) {
          row[targetHeaderMap['費用類型']] = costType;
        }
        if ('費用項目' in targetHeaderMap) {
          row[targetHeaderMap['費用項目']] = '集團費用分攤';
        }
        var hasRowData = false;
        monthColumns.forEach(function (monthCol, monthIndex) {
          if (!(monthCol.header in targetHeaderMap)) {
            return;
          }
          var value = allocations[monthIndex][index];
          if (costType === '人事費用') {
            var directKey =
              config.fields['子公司'] + '|' + config.fields['事業單位'] + '|' + unitKey;
            var directTotals = directPersonnelTotals.get(directKey);
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
  });

  if (outputRows.length === 0) {
    Logger.log('符合條件的資料所有月份皆為 0，僅清除既有資料。');
    return;
  }

  const startRow = targetSheet.getLastRow() + 1;
  targetSheet
    .getRange(startRow, 1, outputRows.length, targetHeaders.length)
    .setValues(outputRows);
}

function accumulateMonthlyTotals_(bucketMap, unitKey, monthColumns, row) {
  if (!bucketMap.has(unitKey)) {
    bucketMap.set(unitKey, monthColumns.map(function () {
      return 0;
    }));
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

function normalizeUnitKey_(value) {
  var normalized = normalizeString_(value);
  if (!normalized) return '';
  normalized = normalized.replace(/：/g, ':');
  var segments = normalized.split(':').map(function (part) {
    return part.trim();
  }).filter(function (part) {
    return part !== '';
  });
  if (segments.length === 0) {
    return '';
  }
  if (segments.length === 1) {
    return segments[0];
  }
  return segments[0] + ' : ' + segments[1];
}

function getGroupAllocationUnitDisplay_(unitKey) {
  var mapping = {
    '執行長室 : 管理中心': '執行長室 : 管理中心 : 管理中心',
    '執行長室 : 財務中心': '執行長室 : 財務中心 : 財務中心',
    '執行長室 : 研發中心': '執行長室 : 研發中心 : 研發中心',
    '執行長室 : 行銷營運中心': '執行長室 : 行銷營運中心 : 行銷營運中心',
    '執行長室 : 客戶價值中心': '執行長室 : 客戶價值中心 : 客戶價值中心',
    '執行長室 : 策略資料中心': '執行長室 : 策略資料中心 : 策略資料中心',
  };
  return mapping[unitKey] || unitKey;
}

function getSplitConfigForUnit_(unitKey) {
  for (var i = 0; i < GROUP_ALLOCATION_CONFIGS_.length; i++) {
    var group = GROUP_ALLOCATION_CONFIGS_[i];
    if (group.unitSet.has(unitKey)) {
      return group.splitConfig;
    }
  }
  return null;
}

function indexOfNth_(str, char, occurrence) {
  var index = -1;
  var fromIndex = 0;
  for (var i = 0; i < occurrence; i++) {
    index = str.indexOf(char, fromIndex);
    if (index === -1) {
      return -1;
    }
    fromIndex = index + 1;
  }
  return index;
}

function assertHeadersPresent_(headerMap, headerNames, contextLabel) {
  headerNames.forEach(function (name) {
    if (!(name in headerMap)) {
      throw new Error((contextLabel || '工作表') + '缺少必要欄位：' + name);
    }
  });
}

function collectDirectPersonnelTotals_(rows, headerMap, monthColumns) {
  var totalsMap = new Map();
  rows.forEach(function (row) {
    var costType = normalizeString_(row[headerMap['費用類型']]);
    if (costType !== '人事費用') {
      return;
    }
    var subsidiary = normalizeString_(row[headerMap['子公司']]);
    if (!subsidiary) {
      return;
    }
    var businessUnit = normalizeString_(row[headerMap['事業單位']]);
    if (businessUnit !== 'POS' && businessUnit !== 'OMO') {
      return;
    }
    var unitKey = normalizeUnitKey_(normalizeString_(row[headerMap['費用單位']]));
    if (!unitKey) {
      return;
    }
    var key = subsidiary + '|' + businessUnit + '|' + unitKey;
    accumulateMonthlyTotals_(totalsMap, key, monthColumns, row);
  });
  return totalsMap;
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
    var value = normalizeString_(row[headerMap[key]]);
    return value === normalizeString_(criteria[key]);
  });
}

function detectMonthColumns_(headers, rows, reservedHeaders) {
  var reserved = reservedHeaders || new Set();
  return headers.reduce(function (acc, header, index) {
    if (!header || reserved.has(header)) {
      return acc;
    }
    if (!isMonthlyHeader_(header)) {
      return acc;
    }
    var hasNumeric = rows.some(function (row) {
      return typeof row[index] === 'number' && !isNaN(row[index]);
    });
    if (hasNumeric) {
      acc.push({ header: header, index: index });
    }
    return acc;
  }, []);
}

function isMonthlyHeader_(header) {
  var MONTH_PATTERN =
    /(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec|Q[1-4]|月|20\d{2}|總計|合計|總額|Total|Sum)/i;
  return MONTH_PATTERN.test(header);
}

function toNumber_(value) {
  if (typeof value === 'number') {
    return value;
  }
  if (typeof value === 'string' && value.trim() !== '') {
    var parsed = parseFloat(value.replace(/,/g, ''));
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

function removeExistingAllocations_(sheet, referenceFields, headerMap) {
  var dataRange = sheet.getDataRange();
  if (dataRange.getNumRows() < 2) {
    return;
  }
  var values = dataRange.getValues();
  var headers = values[0];
  var map = headerMap || buildHeaderMap_(headers);
  var matchFields = referenceFields || {};
  var rowsToDelete = [];
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
