/**
 * Registers the Business Unit specific menus when the spreadsheet opens.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  if (typeof addTaResourceMenus_ === 'function') {
    addTaResourceMenus_(ui);
  } else {
    addTaResourceMenusFallback_(ui);
  }

  if (typeof addMaintenanceSummaryMenu_ === 'function') {
    addMaintenanceSummaryMenu_(ui);
  } else {
    addMaintenanceSummaryMenuFallback_(ui);
  }

  if (typeof addGroupSummaryMenu_ === 'function') {
    addGroupSummaryMenu_(ui);
  } else {
    addGroupSummaryMenuFallback_(ui);
  }
}

function addTaResourceMenusFallback_(ui) {
  (ui || SpreadsheetApp.getUi())
    .createMenu('各專業 TA 專案資源')
    .addItem('人力資源', 'consolidateHeadcount')
    .addItem('所有資源', 'consolidateAllResources')
    .addToUi();
}

function addMaintenanceSummaryMenuFallback_(ui) {
  (ui || SpreadsheetApp.getUi())
    .createMenu('維運費用匯總')
    .addItem('執行匯總', 'consolidateMaintenanceCosts')
    .addToUi();
}

function addGroupSummaryMenuFallback_(ui) {
  (ui || SpreadsheetApp.getUi())
    .createMenu('集團分攤匯總')
    .addItem('執行匯總', 'consolidateGroupAllocations')
    .addToUi();
}
