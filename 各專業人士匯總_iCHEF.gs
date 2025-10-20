const SOURCE_SHEET_NAME = '事業人力資源配置匯總表(自動彙整)';
const SOURCE_COLUMNS = 22; // 欄 A–V

const PERSONNEL_GROUPS = {
  ICHEF: {
    targetSheetName: 'iCHEF_人員總表',
    sources: [
      { id: '1Y1KcaypEZGn_EePnDRYgj12mNw9yEvin7AZRsPhL_7E', label: '策略資料中心' },
      { id: '17a5CYVYBiJgD85EYyV4GGFW4Z8DBbc0mfNcGCyAeml4', label: '管理中心' },
      { id: '1GPAPW3ZM3lY1Qmpb9LocnB27XWJ8AGybGhMNAy2KH0c', label: '財務中心' },
      { id: '1mYeC6DpUqFvnce9leb-098wSYCGqfOR5_HGDM2bcbEc', label: '客戶價值中心' },
      { id: '1Pt2ru3sBxa8TpMfTsfQLjJqdTzZZk0AYpAJvc9TjFY8', label: '行銷營運中心' },
      { id: '1ALyHp6xEMt8xt3T9zpEA52ZUzaDPy-QT-rRh6gU75_g', label: '研發中心' }
    ]
  },
  QLR: {
    targetSheetName: 'QLR_人員總表',
    sources: [
      { id: '10jcSlS4RvuGm1DK4KnyKz0YBVxHE_2OCnrBgAgqui9U', label: 'QLR 管理中心' },
      { id: '1X6c37n6s4XQumB4mD-N_M_Pqv4ao9f9S7Zea6c6TXRI', label: 'QLR 研發中心' },
      { id: '1-jof2z4-D2KbMRq2F7h0BW_p7cQy95Ppb-R1k9b6oeQ', label: 'QLR 行銷營運中心' }
    ]
  }
};

function syncPersonnelData(groupKey = 'ICHEF') {
  const group = PERSONNEL_GROUPS[groupKey];
  if (!group) throw new Error(`找不到分組設定：${groupKey}`);

  const targetSheet = SpreadsheetApp.getActive().getSheetByName(group.targetSheetName);
  if (!targetSheet) throw new Error(`找不到目標分頁：${group.targetSheetName}`);

  const aggregatedValues = [];

  group.sources.forEach((source, index) => {
    const sourceSheet = SpreadsheetApp.openById(source.id).getSheetByName(SOURCE_SHEET_NAME);
    if (!sourceSheet) {
      throw new Error(`找不到來源分頁：${SOURCE_SHEET_NAME}（${source.label}）`);
    }

    const lastRow = sourceSheet.getLastRow();
    if (lastRow === 0) return; // 跳過空表

    const values = sourceSheet.getRange(1, 1, lastRow, SOURCE_COLUMNS).getValues();
    if (values.length === 0) return;

    if (aggregatedValues.length === 0) {
      aggregatedValues.push(...values);
    } else {
      aggregatedValues.push(...values.slice(1)); // 後續來源避開標題列
    }
  });

  targetSheet.getRange('A:V').clearContent();

  if (aggregatedValues.length === 0) return;

  targetSheet
    .getRange(1, 1, aggregatedValues.length, SOURCE_COLUMNS)
    .setValues(aggregatedValues);
}

function consolidateHeadcount() {
  syncPersonnelData('ICHEF');
}

function consolidateHeadcountQLR() {
  syncPersonnelData('QLR');
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('匯總人事')
    .addItem('匯總 iCHEF', 'consolidateHeadcount')
    .addItem('匯總 QLR', 'consolidateHeadcountQLR')
    .addToUi();
}
