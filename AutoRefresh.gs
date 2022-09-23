function autoRefresh() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Dashboard");
  sheet.getRange("B1").setValue(new Date().toTimeString());
}
