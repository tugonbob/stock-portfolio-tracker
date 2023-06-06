function GetRecentSuperInvestorTrades() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Super Investors");
  let trades = sheet.getRange("H2:H").getValues();

  let recentUniqueTrades = [];

  for (let i = 0; i < trades.length; i++) {
    if (trades[i][0] === "") break; // skip empty rows

    let dateIndex = 0;
    let tickerIndex = -1;
    let firstCapitalLetter = true; // to get the second capital letter
    for (let j = 0; j < trades[i][0].length; j++) {
      // if isLetter and isUppercase
      if (
        trades[i][0].charAt(j).match(/[a-z]/i) &&
        trades[i][0].charAt(j) === trades[i][0].charAt(j).toUpperCase()
      ) {
        // if is the 2nd capital letter found
        if (!firstCapitalLetter) {
          tickerIndex = j;
          break;
        }
        firstCapitalLetter = false;
      }
    }

    let splitData = trades[i][0].split(" - ");
    let nameIndex = 0;
    let sharesIndex = -1;
    for (let j = splitData[1].length - 1; j >= 0; j--) {
      // if is letter
      if (splitData[1].charAt(j).match(/[a-z]/i)) {
        sharesIndex = j + 1;
        break;
      }
    }

    let priceIndex = -1;
    // find first comma starting from the back to get price
    for (let j = splitData[1].length - 1; j >= 0; j--) {
      if (splitData[1].charAt(j) === ",") {
        priceIndex = j + 4;
        break;
      }
    }

    let date = splitData[0].substring(dateIndex, tickerIndex);
    let ticker = splitData[0].substring(tickerIndex);
    let name = splitData[1].substring(nameIndex, sharesIndex);
    let shares = parseInt(
      splitData[1]
        .substring(sharesIndex, priceIndex)
        .replace(/,/g, "")
        .replace(/\./g, "")
    );
    let price = parseFloat(splitData[1].substring(priceIndex));

    // Lowercase the whole name except first letter of each word
    let nameSplit = name.toLowerCase().split(" ");
    for (j = 0; j < nameSplit.length; j++) {
      nameSplit[j] =
        nameSplit[j].charAt(0).toUpperCase() + nameSplit[j].slice(1);
    }
    name = nameSplit.join(" ");

    // find if name already exists in recentUniqueTrades arr
    let index = -1;
    for (let j = 0; j < recentUniqueTrades.length; j++) {
      if (recentUniqueTrades[j][1] === name) {
        index = j;
        break;
      }
    }

    // if name already exists, do weighted average
    if (index > -1) {
      let previousShares = parseInt(recentUniqueTrades[index][2]);
      let previousPrice = parseFloat(recentUniqueTrades[index][3]);

      previousPrice =
        (shares * price + previousShares * previousPrice) /
        (previousShares + shares);
      shares += previousShares;

      recentUniqueTrades[index][3] = previousPrice;
      recentUniqueTrades[index][2] = shares;
    }

    // add top 9 unique trades to recentUniqueTrades arr
    if (index < 0 && recentUniqueTrades.length < 10)
      recentUniqueTrades.push([date + " - " + ticker, name, shares, price]);

    if (recentUniqueTrades.length === 10) {
      break;
    }
  }
  return recentUniqueTrades;
}
