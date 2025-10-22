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
      { id: '1ALyHp6xEMt8xt3T9zpEA52ZUzaDPy-QT-rRh6gU75_g', label: '研發中心' },
      { id: '1Nk_TXUKf5uX0P-fmbP9jlSio6qB1pH1V4XFFsUWWIzM', label: '執行長室' },
      { id: '1XCORsakXwHErz2AxHTTBY1a1yxsN78scGKOd7_uIPss', label: '執行長室 OMO ' }
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

  let headerRow = null;
  const dataRows = [];

  group.sources.forEach((source, index) => {
    const sourceSheet = SpreadsheetApp.openById(source.id).getSheetByName(SOURCE_SHEET_NAME);
    if (!sourceSheet) {
      throw new Error(`找不到來源分頁：${SOURCE_SHEET_NAME}（${source.label}）`);
    }

    const lastRow = sourceSheet.getLastRow();
    if (lastRow === 0) return; // 跳過空表

    const values = sourceSheet.getRange(1, 1, lastRow, SOURCE_COLUMNS).getValues();
    if (values.length === 0) return;

    const [currentHeader, ...currentData] = values;

    if (!headerRow) {
      headerRow = currentHeader;
    }

    if (!headerRow) {
      headerRow = currentHeader;
    }

    const rowsToAppend = headerRow ? currentData : values;
    dataRows.push(...rowsToAppend);
  });

  targetSheet.getRange('A:V').clearContent();

  if (!headerRow) return;

  const output = [headerRow, ...dataRows];

  targetSheet
    .getRange(1, 1, output.length, SOURCE_COLUMNS)
    .setValues(output);
}

function consolidateHeadcount() {
  syncPersonnelData('ICHEF');
}

function consolidateHeadcountQLR() {
  syncPersonnelData('QLR');
}

function onOpen(e) {
  const ui = (e && e.source && typeof e.source.getUi === 'function')
    ? e.source.getUi()
    : SpreadsheetApp.getUi();
  addPersonnelSyncMenus_(ui);
}

function addPersonnelSyncMenus_(ui) {
  (ui || SpreadsheetApp.getUi())
    .createMenu('匯總人事')
    .addItem('匯總 iCHEF', 'consolidateHeadcount')
    .addItem('匯總 QLR', 'consolidateHeadcountQLR')
    .addToUi();
}
