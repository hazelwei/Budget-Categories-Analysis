function addMaintenanceSummaryMenu_(ui) {
  (ui || SpreadsheetApp.getUi())
    .createMenu('維運費用匯總')
    .addItem('執行匯總', 'consolidateMaintenanceCosts')
    .addToUi();
}

function consolidateMaintenanceCosts() {
  const SOURCE_SHEET_NAME = '維運攤分';
  const TARGET_SHEET_NAME = '2.7_維運所需資源（清單）';

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
    businessUnit: 'OMO',
    siExclusions: ['Maintenance', 'Corporation'],
    taExclusions: ['Baseline', 'Corporation'],
    projectCodeExclusions: ['NA']
  };

  let header = null;
  const aggregatedRows = [];

  SOURCE_SPREADSHEET_IDS.forEach(spreadsheetId => {
    const dataset = loadSourceDataset_(spreadsheetId, SOURCE_SHEET_NAME);
    if (!dataset.header.length) return;

    if (!header) header = dataset.header;

    const effectiveRows = dataset.rows
      .filter(row => rowMatchesFilters_(row, filters))
      .filter(hasMaintenanceContent_);

    aggregatedRows.push(...effectiveRows);
  });

  if (!header || !header.length) {
    throw new Error('來源試算表沒有可用的標題列');
  }

  const targetSpreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  writeToTarget_(targetSpreadsheetId, TARGET_SHEET_NAME, header, aggregatedRows);
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
