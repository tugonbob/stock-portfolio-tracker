function portfolioValueTracker() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Chart Data");
  let currentPortfolioValue = sheet.getRange("A4").getValue();
  let portfolioValueHistory = sheet.getRange("E2:E").getValues();
  const lastRow =
    portfolioValueHistory.findIndex((element) => element[0] === "") + 2;
  sheet.getRange(`E${lastRow}`).setValue(currentPortfolioValue);
  sheet
    .getRange(`D${lastRow}`)
    .setValue(Utilities.formatDate(new Date(), "GMT-6", "MM/dd/yyyy"));
}

function minuteGainsTracker() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Chart Data");
  let currentGains = sheet.getRange("A2").getValue();
  let gainsHistory = sheet.getRange("G2:G").getValues();
  const lastRow = gainsHistory.findIndex((element) => element[0] === "") + 2;

  let now = new Date();
  let fivePm = new Date().setHours(17, 0, 0);
  let eightAm = new Date().setHours(8, 0, 0);

  if (now < fivePm && now > eightAm) {
    sheet.getRange(`F${lastRow}`).setValue(new Date().toLocaleTimeString());
    sheet.getRange(`G${lastRow}`).setValue(currentGains);
  }
}

function deleteMinuteGains() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Chart Data");
  sheet.getRange("G2:G").clear();
  sheet.getRange("F2:F").clear();
}

function dailyGainsTracker() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Chart Data");
  let currentGains = sheet.getRange("A2").getValue();
  let gainsHistory = sheet.getRange("I2:I").getValues();
  const lastRow = gainsHistory.findIndex((element) => element[0] === "") + 2;
  sheet.getRange(`I${lastRow}`).setValue(currentGains);
  sheet
    .getRange(`H${lastRow}`)
    .setValue(Utilities.formatDate(new Date(), "GMT-6", "MM/dd/yyyy"));
}

function GetDynamicChart(timePeriod) {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Chart Data");
  let dailyGains = sheet.getRange("H2:I").getValues();
  const lastRow = dailyGains.findIndex((element) => element[0] === "") + 2;
  dailyGains = dailyGains.slice(0, lastRow);

  if (timePeriod === "1M") {
    if (lastRow < 32) return sheet.getRange(`H2:I${lastRow}`).getValues();
    else return sheet.getRange(`H${lastRow - 30}:I${lastRow}`).getValues();
  } else if (timePeriod === "YTD") {
    let now = new Date();
    let start = new Date(now.getFullYear(), 0, 0);
    let day = Math.floor((now - start) / (1000 * 60 * 60 * 24));

    if (lastRow < day + 1) return sheet.getRange(`H2:I${lastRow}`).getValues();
    else return sheet.getRange(`H${lastRow - day}:I${lastRow}`).getValues();
  } else if (timePeriod === "1Y") {
    if (lastRow < 367) return sheet.getRange(`H2:I${lastRow}`).getValues();
    else;
    return sheet.getRange(`H${lastRow - 365}:I${lastRow}`).getValues();
  } else if (timePeriod === "5Y") {
    if (lastRow - 1828) return sheet.getRange(`H2:I${lastRow}`).getValues();
    else return sheet.getRange(`H${lastRow - 1826}:I${lastRow}`).getValues();
  } else {
    return sheet.getRange(`H2:I${lastRow}`).getValues();
  }
}
