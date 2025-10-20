/**
 * Consolidated onOpen handler to register all custom menus.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  if (typeof addHeadcountMenu_ === 'function') {
    addHeadcountMenu_(ui);
  }
  if (typeof addPersonnelSyncMenus_ === 'function') {
    addPersonnelSyncMenus_(ui);
  }
  if (typeof addTaResourceMenus_ === 'function') {
    addTaResourceMenus_(ui);
  }
  if (typeof addMaintenanceMenu_ === 'function') {
    addMaintenanceMenu_(ui);
  }
}
