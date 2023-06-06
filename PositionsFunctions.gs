function GetUniqueTickers() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("History");
  let tickers = sheet.getRange("B3:B").getValues();

  let uniqueTickers = [];
  for (let i = 0; i < tickers.length; i++) {
    if (tickers[i][0] === "Cash") continue;
    if (!uniqueTickers.includes(tickers[i][0]))
      uniqueTickers.push(tickers[i][0]);
  }
  uniqueTickers.pop();
  return uniqueTickers;
}

function GetOwnedShareNumber(ticker) {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("History");
  let tickers = sheet.getRange("B3:B").getValues();
  let actions = sheet.getRange("C3:C").getValues();
  let quantities = sheet.getRange("D3:D").getValues();

  let sum = 0;
  for (let i = 0; i < tickers.length; i++) {
    if (tickers[i][0] !== ticker) continue;
    if (actions[i][0] === "Dividend") continue;

    sum += quantities[i][0];
  }
  return sum;
}

function GetShareMaturity(ticker) {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("History");
  let dates = sheet.getRange("A3:A").getValues();
  let tickers = sheet.getRange("B3:B").getValues();
  let actions = sheet.getRange("C3:C").getValues();
  let quantities = sheet.getRange("D3:D").getValues();

  let sum = 0;
  let prevYear = new Date(new Date().setFullYear(new Date().getFullYear() - 1));
  for (let i = 0; i < tickers.length; i++) {
    if (tickers[i][0] !== ticker) continue;
    if (actions[i][0] == "Dividend") continue;

    if (dates[i][0] > prevYear) {
      let diffInMs = new Date() - dates[i][0];
      let diffInDays = diffInMs / (1000 * 60 * 60 * 24);

      sum += quantities[i][0] * (diffInDays / 365);
    } else {
      sum += quantities[i][0];
    }
  }
  return sum;
}

function GetTotalDividends(ticker) {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("History");
  let tickers = sheet.getRange("B3:B").getValues();
  let actions = sheet.getRange("C3:C").getValues();
  let prices = sheet.getRange("E3:E").getValues();

  let sum = 0;
  for (let i = 0; i < tickers.length; i++) {
    if (tickers[i][0] !== ticker) continue;
    if (actions[i][0] !== "Dividend") continue;

    sum += prices[i][0];
  }
  return sum;
}

function GetAveragePurchasePrice(ticker) {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("History");
  let history = sheet.getRange("A3:E").getValues();

  let queue = []; // for FIFO
  for (let i = 0; i < history.length; i++) {
    if (history[i][1] !== ticker) continue;
    if (history[i][2] !== "Buy" && history[i][2] !== "Sell") continue;

    if (history[i][2] === "Buy") {
      for (let j = 0; j < history[i][3]; j++) {
        queue.push(history[i][4]);
      }
    } else if (history[i][2] === "Sell") {
      for (let j = 0; j < -1 * history[i][3]; j++) {
        queue.shift();
      }
    }
  }

  let averagePrice = 0;
  for (let i = 0; i < queue.length; i++) {
    averagePrice += queue[i];
  }
  return averagePrice / queue.length;
}

function GetRealizedGain(ticker) {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("History");
  let history = sheet.getRange("A3:E").getValues();

  let queue = []; // for FIFO
  let realizedGains = 0;
  for (let i = 0; i < history.length; i++) {
    if (history[i][1] !== ticker) continue;
    if (history[i][2] !== "Buy" && history[i][2] !== "Sell") continue;

    if (history[i][2] === "Buy") {
      for (let j = 0; j < history[i][3]; j++) {
        queue.push(history[i][4]);
      }
    } else if (history[i][2] === "Sell") {
      for (let j = 0; j < -1 * history[i][3]; j++) {
        realizedGains += history[i][4] - queue.shift();
      }
    }
  }
  return realizedGains;
}

function GetYtdStats() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let historySheet = ss.getSheetByName("History");
  let chartDataSheet = ss.getSheetByName("Chart Data");
  let history = historySheet.getRange("A3:F").getValues();
  let dailyGains = chartDataSheet.getRange("I2:J").getValues();
  let portfolioValue = chartDataSheet.getRange("E2:F").getValues();

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

function GetLastPortfolioValue() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let chartDataSheet = ss.getSheetByName("Chart Data");
  let portfolioValue = chartDataSheet.getRange("E2:F").getValues();

  let lastRow = portfolioValue.findIndex((element) => element[0] === "") - 1;

  return portfolioValue[lastRow][1];
}

function GetBenchmarchSp500ShareNumberAndAveragePrice() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let historySheet = ss.getSheetByName("History");
  let vooSheet = ss.getSheetByName("VOO Price History");
  let history = historySheet.getRange("A3:F").getValues();
  let voo = vooSheet.getRange("A2:B").getValues();

  let lastRow = history.findIndex((element) => element[0] === "") - 1;

  function findVooPriceByDate(date) {
    for (let j = 1; j < voo.length; j++) {
      date.setHours(0, 0, 0, 0);
      let vooDate = voo[j][0];
      vooDate.setHours(0, 0, 0, 0);
      if (vooDate == date) {
        return parseFloat(voo[j][1]);
      } else if (voo[j][0] > date) {
        return parseFloat(voo[j - 1][1]);
      }
    }
  }

  let sharePrices = []; // FIFO queue
  let cash = 0;
  for (let i = 0; i < lastRow; i++) {
    if (history[i][2] !== "Deposit" && history[i][2] !== "Withdrawal") continue;

    let vooPrice = findVooPriceByDate(history[i][0]);
    let amount = parseFloat(history[i][5]);
    if (history[i][2] === "Deposit") {
      amount += cash;
      cash = 0;
      let sharesAfford = Math.floor(amount / vooPrice);
      let remainder = (amount % vooPrice) / vooPrice;
      cash += vooPrice * remainder;
      for (let z = 0; z < sharesAfford; z++) {
        sharePrices.push(vooPrice);
      }
    } else {
      amount -= cash;
      cash = 0;
      let sharesNeeded = Math.floor(amount / vooPrice) + 1;
      let remainder = (amount % vooPrice) / vooPrice;
      cash += (1 - remainder) * vooPrice;
      for (let z = 0; z < sharesNeeded; z++) {
        sharePrices.shift();
      }
    }
  }

  let numberShares = sharePrices.length;

  if (numberShares === 0) return [[0, 0]];

  let sum = sharePrices.reduce((a, b) => a + b);
  let averagePrice = sum / numberShares;

  return [[numberShares, averagePrice]];
}
