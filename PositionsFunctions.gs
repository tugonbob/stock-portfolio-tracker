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
