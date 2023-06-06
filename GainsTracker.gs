function portfolioValueTracker() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Chart Data");
  let currentPortfolioValue = sheet.getRange("A4").getValue();
  let portfolioValueHistory = sheet.getRange("F2:F").getValues();
  const lastRow =
    portfolioValueHistory.findIndex((element) => element[0] === "") + 2;
  sheet.getRange(`F${lastRow}`).setValue(currentPortfolioValue);
  sheet
    .getRange(`E${lastRow}`)
    .setValue(Utilities.formatDate(new Date(), "GMT-6", "MM/dd/yyyy"));
}

function minuteGainsTracker() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Chart Data");
  let currentGains = sheet.getRange("A2").getValue();
  let gainsHistory = sheet.getRange("H2:H").getValues();
  const lastRow = gainsHistory.findIndex((element) => element[0] === "") + 2;

  let now = new Date();
  let fivePm = new Date().setHours(17, 0, 0);
  let eightAm = new Date().setHours(8, 0, 0);

  if (now < fivePm && now > eightAm) {
    sheet.getRange(`G${lastRow}`).setValue(new Date().toLocaleTimeString());
    sheet.getRange(`H${lastRow}`).setValue(currentGains);
  }
}

function deleteMinuteGains() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Chart Data");
  sheet.getRange("H2:H").clear();
  sheet.getRange("G2:G").clear();
}

function dailyGainsTracker() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Chart Data");
  let currentGains = sheet.getRange("A2").getValue();
  let gainsHistory = sheet.getRange("J2:J").getValues();
  const lastRow = gainsHistory.findIndex((element) => element[0] === "") + 2;
  sheet.getRange(`J${lastRow}`).setValue(currentGains);
  sheet
    .getRange(`I${lastRow}`)
    .setValue(Utilities.formatDate(new Date(), "GMT-6", "MM/dd/yyyy"));
}

function sp500GainsTracker() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let chartDataSheet = ss.getSheetByName("Chart Data");
  let positionsSheet = ss.getSheetByName("Positions");
  let sp500Gains = positionsSheet.getRange("K3").getValue();
  let gainsHistory = chartDataSheet.getRange("K2:K").getValues();
  const lastRow = gainsHistory.findIndex((element) => element[0] === "") + 2;
  chartDataSheet.getRange(`K${lastRow}`).setValue(sp500Gains);
}

function GetDynamicChart(timePeriod) {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Chart Data");
  let dailyGains = sheet.getRange("I2:J").getValues();

  const lastRow = dailyGains.findIndex((element) => element[0] === "") + 2;
  dailyGains = dailyGains.slice(0, lastRow);

  if (timePeriod === "1M") {
    if (lastRow < 32) return sheet.getRange(`I2:K${lastRow}`).getValues();
    else return sheet.getRange(`I${lastRow - 30}:K${lastRow}`).getValues();
  } else if (timePeriod === "YTD") {
    let now = new Date();
    let start = new Date(now.getFullYear(), 0, 0);
    let day = Math.floor((now - start) / (1000 * 60 * 60 * 24));

    if (lastRow < day + 1) return sheet.getRange(`I2:K${lastRow}`).getValues();
    else return sheet.getRange(`I${lastRow - day}:K${lastRow}`).getValues();
  } else if (timePeriod === "1Y") {
    if (lastRow < 367) return sheet.getRange(`I2:K${lastRow}`).getValues();
    else;
    return sheet.getRange(`I${lastRow - 365}:K${lastRow}`).getValues();
  } else if (timePeriod === "5Y") {
    if (lastRow - 1828) return sheet.getRange(`I2:K${lastRow}`).getValues();
    else return sheet.getRange(`I${lastRow - 1826}:K${lastRow}`).getValues();
  } else {
    return sheet.getRange(`G2:H${lastRow}`).getValues();
  }
}
