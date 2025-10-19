function syncPersonnelData() {
  const SOURCE_SPREADSHEET_ID = '1Y1KcaypEZGn_EePnDRYgj12mNw9yEvin7AZRsPhL_7E';
  const SOURCE_SHEET_NAME = '事業人力資源配置匯總表(中心主管填)';
  const TARGET_SHEET_NAME = 'iCHEF_人員總表';
  const SOURCE_COLUMNS = 22; // 欄 A–V

  const sourceSheet = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID).getSheetByName(SOURCE_SHEET_NAME);
  if (!sourceSheet) throw new Error(`找不到來源分頁：${SOURCE_SHEET_NAME}`);

  const lastRow = sourceSheet.getLastRow();
  if (lastRow === 0) {
    SpreadsheetApp.getActive().getSheetByName(TARGET_SHEET_NAME).clearContents();
    return;
  }

  const values = sourceSheet.getRange(1, 1, lastRow, SOURCE_COLUMNS).getValues();

  const targetSheet = SpreadsheetApp.getActive().getSheetByName(TARGET_SHEET_NAME);
  if (!targetSheet) throw new Error(`找不到目標分頁：${TARGET_SHEET_NAME}`);

  targetSheet.clearContents();
  targetSheet.getRange(1, 1, values.length, values[0].length).setValues(values);
}
