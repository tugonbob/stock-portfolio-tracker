function GetLatestSecFilings() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let chartDataSheet = ss.getSheetByName("Positions");
  let tickers = chartDataSheet.getRange("A4:A").getValues();
  let filingDates = chartDataSheet.getRange("N4:Q").getValues();

  console.log(filingDates);

  let result = [];
  for (let i = 0; i < filingDates.length; i++) {
    for (let j = 0; j < filingDates[i].length; j++) {
      let dateStr = "";
      let filingType = "";
      if (filingDates[i][j].charAt(0) === "A") {
        dateStr = filingDates[i][j + 1];
        filingType = "Annual";
      } else if (filingDates[i][j].charAt(0) === "1") {
        dateStr = filingDates[i][j + 1];
        filingType = "Quarter";
      }

      let index = dateStr.indexOf("-");
      if (index > -1) {
        let formattedDateStr = new Date(
          dateStr.substring(index + 2, index + 12)
        ).toLocaleDateString("en-US", {
          month: "short",
          day: "numeric",
          year: "numeric",
        });
        result.push([tickers[i][0], filingType, formattedDateStr]);
      }
    }
  }

  result.sort((a, b) => {
    if (new Date(a[2]) === new Date(b[2])) return 0;
    else return new Date(a[2]) > new Date(b[2]) ? -1 : 1;
  });

  return result.slice(0, 10);
}
