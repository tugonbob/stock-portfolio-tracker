# Portfolio Tracker

## Link to Spreadsheet. Take a look!

https://docs.google.com/spreadsheets/d/1VVdyaR6301wR6bW8tCLX5kh-q9Qt_3s0AiSrbTvzYKQ/edit?usp=sharing

# Tab Descriptions

### Dashboard

Displays my portfolio balances, some useful portfolio stats, a gains chart, a map of the S&P 500, a holdings pie chart, some recent trades my by respected investors, top buys last quarter of respected investors, recent annual and quarterly reports of companies that I own, and daily vs YTD performance of major indices

### History

A history of every trade, dividend, despoit, and withdrawal that I've ever done

### Positions

Some stats on every ticker that I've traded on before

### Chart Data

Used to keep track of daily portfolio value, gains minute by minute, and daily portfolio gains. Used to be displayed on gains chart on "Dashboard" tab and calculate some stats

### Market Map

Used to scrape data for the S&P500 map on "Dashboard" tab. Also keeps track of daily and YTD percentage changes of major indices

### Super Investors

Used to scrape data from https://www.dataroma.com/m/home.php and to display that data on the "Dashboard" tab

# Script Descriptions

## PositionsFunctions.gs

### GetUniqueTickers()

Look through "History" tab and list all tickers once in the "Positions" tab

### GetOwnedShareNumber(ticker)

Given a ticker, find the amount of shares currently owned. It does this by looking through "History" tab and adding to the total when shares are bought and subtracting when shares are sold

### GetTotalDividends(ticker)

Given a ticker, add up total dividends made with this ticker by looking through "History" tab

### GetAveragePurchasePrice(ticker)

Given a ticker, calculate the average purchase price of the shares that are currently owned. I loop through the "History" tab and use a queue data structure to keep track of which shares have been sold or are currently held and their purchase prices

### GetRealizedGain(ticker)

Calculate the total realized gains of all shares ever traded with the given ticker

## DashboardFunctions.gs

### GetYtdStats()

Calculate YTD Invested, YTD Dividends, YTD Gains, and YTD percentage gain

### GetPortfolioCagr()

Calculate total portfolio CAGR using the time weighted return method: https://www.fool.com/about/how-to-calculate-investment-returns/

### GetLastPortfolioValue()

Get the most updated total portfolio value

## Gains Tracker

### portfolioValueTracker()

Record my daily total portfolio value to the "Chart Data" tab

### minuteGainsTracker()

Every minute, save my total gains to "Chart Data" tab during market open times

### deleteMinuteGains()

Everyday before market open, delete yesterday's minute gains

### dailyGainsTracker()

Everyday, record total gains after market close

### GetDynamicChart(timeperiod)

Given a time period of either 1D, 1M, YTD, 1Y, 5Y, and Lifetime, get the data points for these time periods to be displayed on the Performance chart in the "Dashboard" tab

## SuperInvestor.gs

### GetRecentSuperInvestorTrades()

Datascrape recent data "Super Investor" trades from https://www.dataroma.com/m/home.php

## GetLastSecFilings.gs

### GetLatestSecFilings()

Find the latest annual and quarterly reports of the companies that I own, and display them on "Dashboard" tab

## AutoRefresh.gs

### autoRefresh()

Change the refresh number every time someone opens the Google Sheet. This is used to update the data scrapers
