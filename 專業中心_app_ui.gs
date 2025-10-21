/**
 * Consolidated onOpen handler to register all custom menus.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  if (typeof addHeadcountMenu_ === 'function') {
    addHeadcountMenu_(ui);
  }
  if (typeof addMaintenanceMenu_ === 'function') {
    addMaintenanceMenu_(ui);
  }
  if (typeof addGroupAllocationMenu_ === 'function') {
    addGroupAllocationMenu_(ui);
  }
}
