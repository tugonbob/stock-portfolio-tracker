function GetYtdStats() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let historySheet = ss.getSheetByName("History");
  let chartDataSheet = ss.getSheetByName("Chart Data");
  let history = historySheet.getRange("A3:F").getValues();
  let dailyGains = chartDataSheet.getRange("H2:I").getValues();
  let portfolioValue = chartDataSheet.getRange("D2:E").getValues();

  let ytdStats = [];
  let ytdInvested = 0;
  let ytdDividends = 0;
  let ytdDollarChange = 0;
  let ytdPercentChange = 0;
  let ttmDividends = 0;
  let ttmHourlySalary = 0;

  // get most recent december 31st
  const dec31 = new Date(new Date().getFullYear() - 1, 11, 31).setHours(
    0,
    0,
    0,
    0
  );
  const oneYearAgo = new Date(
    new Date().setFullYear(new Date().getFullYear() - 1)
  );

  // get ytdInvested and ytdDividends and ttmDividends
  for (let i = 0; i < history.length; i++) {
    if (history[i][0] > dec31) {
      if (history[i][2] === "Deposit") ytdInvested += history[i][4];
      else if (history[i][2] === "Withdrawal") ytdInvested -= history[i][4];
      else if (history[i][2] === "Dividend") ytdDividends += history[i][4];
    }

    if (history[i][0] > oneYearAgo) {
      if (history[i][2] === "Dividend") ttmDividends += history[i][4];
    }
  }

  //get last non-empty row
  const lastRow = dailyGains.findIndex((element) => element[0] === "") - 1;
  // find the index of the row that is most recent dec 31
  const dec31Index = dailyGains.findIndex(
    (element) => element[0].setHours(0, 0, 0, 0) === dec31
  );

  // get ytdDollarChange and ytdPercentChange
  ytdDollarChange = dailyGains[lastRow][1] - dailyGains[dec31Index][1];
  ytdPercentChange = ytdDollarChange / portfolioValue[lastRow][1];

  // get number of work days in a year
  let workDays = 365 * (5 / 7);

  // get ytdHourlySalary
  ttmHourlySalary = (ytdDollarChange + ttmDividends) / workDays;

  ytdStats.push([ytdInvested]);
  ytdStats.push([ytdDividends]);
  ytdStats.push([ytdDollarChange]);
  ytdStats.push([ytdPercentChange]);
  ytdStats.push([ttmHourlySalary]);

  return ytdStats;
}

function GetPortfolioCagr() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let historySheet = ss.getSheetByName("History");
  let chartDataSheet = ss.getSheetByName("Chart Data");
  let history = historySheet.getRange("A3:F").getValues();
  let portfolioValue = chartDataSheet.getRange("D2:E").getValues();

  let hpr = [];

  let endingValue = 0;
  for (let i = 0; i < history.length; i++) {
    if (history[i][2] !== "Deposit" && history[i][2] !== "Withdrawal") continue;

    let valueAfterCashFlow;
    if (history[i][2] === "Deposit")
      valueAfterCashFlow = endingValue + history[i][5];
    else if (history[i][2] === "Withdrawal")
      valueAfterCashFlow = endingValue - history[i][5];

    let nextCashFlowIndex;
    for (let j = i + 1; j < history.length; j++) {
      if (history[j][2] === "Withdrawal" || history[j][2] === "Deposit") {
        nextCashFlowIndex = j;
        break;
      }
    }

    if (nextCashFlowIndex) {
      let nextCashFlowDate = history[nextCashFlowIndex][0];
      for (let j = 0; j < portfolioValue.length; j++) {
        if (portfolioValue[j][0] <= nextCashFlowDate)
          endingValue = portfolioValue[j][1];
        else break;
      }
    } else {
      let lastRow =
        portfolioValue.findIndex((element) => element[0] === "") - 1;
      endingValue = portfolioValue[lastRow][1];
    }

    hpr.push(endingValue / valueAfterCashFlow - 1);
  }

  let twr = 1;
  for (let i = 0; i < hpr.length; i++) {
    twr *= 1 + hpr[i];
  }

  twr = twr - 1;

  let start = new Date("April 18, 2017 00:00:00");
  let ageDifMs = Date.now() - start;
  let ageDate = new Date(ageDifMs); // miliseconds from epoch
  let n = Math.abs(ageDate.getUTCFullYear() - 1970);

  let cagr = Math.pow(1 + twr, 1 / n) - 1;
  return cagr;
}

function GetLastPortfolioValue() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let chartDataSheet = ss.getSheetByName("Chart Data");
  let portfolioValue = chartDataSheet.getRange("D2:E").getValues();

  let lastRow = portfolioValue.findIndex((element) => element[0] === "") - 1;

  return portfolioValue[lastRow][1];
}
