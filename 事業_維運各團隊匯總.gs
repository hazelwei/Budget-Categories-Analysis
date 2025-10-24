function addMaintenanceSummaryMenu_(ui) {
  (ui || SpreadsheetApp.getUi())
    .createMenu('維運費用匯總')
    .addItem('執行匯總', 'consolidateMaintenanceCosts')
    .addToUi();
}

function consolidateMaintenanceCosts() {
    // TODO: 使用前請更新對應的 ALLOWED_SUBSIDIARIES,ALLOWED_BUSINESS_UNITS。
  const SOURCE_SHEET_NAME = '維運攤分';
  const TARGET_SHEET_NAME = '2.7_維運所需資源（清單）';
  const ALLOWED_SUBSIDIARIES = new Set(['TW']);
  const ALLOWED_BUSINESS_UNITS = new Set(['OMO', '2C']);

  const SOURCE_SPREADSHEET_IDS = [
      // TODO: 使用前請更新對應的 SPREADSHEET_ID。
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

  let header = null;
  let columnIndex = null;
  const aggregatedRows = [];

  SOURCE_SPREADSHEET_IDS.forEach(spreadsheetId => {
    const dataset = loadMaintenanceDataset_(spreadsheetId, SOURCE_SHEET_NAME);
    if (!dataset.header.length) return;

    if (!header) {
      header = dataset.header;
      columnIndex = resolveMaintenanceColumnIndex_(header);
    }

    const effectiveRows = dataset.rows
      .filter(row => isAllowedMaintenanceRow_(row, columnIndex, ALLOWED_SUBSIDIARIES, ALLOWED_BUSINESS_UNITS))
      .filter(hasMaintenanceContent_);

    aggregatedRows.push(...effectiveRows);
  });

  if (!header || !header.length) {
    throw new Error('來源試算表沒有可用的標題列');
  }

  const targetSpreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  writeToTarget_(targetSpreadsheetId, TARGET_SHEET_NAME, header, aggregatedRows);
}

function loadMaintenanceDataset_(spreadsheetId, sheetName) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log(`跳過來源：${spreadsheetId}，沒有找到分頁：${sheetName}`);
    return { header: [], rows: [] };
  }
  return readSheetDataset_(sheet);
}

function hasMaintenanceContent_(row) {
  return row.some(value => {
    if (value === null || value === undefined) return false;
    if (value instanceof Date) return true;
    if (typeof value === 'number') return value !== 0;
    if (typeof value === 'boolean') return true;
    if (typeof value === 'string') {
      return maintenanceNormalizeString_(value) !== '';
    }
    return maintenanceNormalizeString_(value) !== '';
  });
}

function maintenanceNormalizeString_(value) {
  if (typeof normalizeString_ === 'function') {
    return normalizeString_(value);
  }
  if (value === null || value === undefined) return '';
  if (typeof value === 'string') return value.trim();
  return String(value).trim();
}

function resolveMaintenanceColumnIndex_(header) {
  const subsidiary = header.indexOf('子公司');
  const businessUnit = header.indexOf('事業單位');
  if (subsidiary === -1 || businessUnit === -1) {
    throw new Error('來源分頁缺少必要欄位：子公司 或 事業單位');
  }
  return { subsidiary, businessUnit };
}

function isAllowedMaintenanceRow_(row, columnIndex, subsidiaryWhitelist, businessUnitWhitelist) {
  if (!columnIndex) return true;
  const normalize = maintenanceNormalizeString_;

  if (columnIndex.subsidiary > -1 && subsidiaryWhitelist && subsidiaryWhitelist.size) {
    const subsidiary = normalize(row[columnIndex.subsidiary]);
    if (subsidiary && !subsidiaryWhitelist.has(subsidiary)) {
      return false;
    }
  }

  if (columnIndex.businessUnit > -1 && businessUnitWhitelist && businessUnitWhitelist.size) {
    const businessUnit = normalize(row[columnIndex.businessUnit]);
    if (businessUnit && !businessUnitWhitelist.has(businessUnit)) {
      return false;
    }
  }

  return true;
}
