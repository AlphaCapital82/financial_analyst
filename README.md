What This Tool Does

This Python tool helps you analyze a stock by connecting to a Google Sheet, performing calculations, and showing the results in a simple desktop window.

How It Works

1. Connects to Google Sheets
   It uses a secure login file to access your spreadsheet and allows the script to read and write data.

2. The spreadsheet itself is connected to a financial database
   The Google Sheet automatically pulls financial numbers (such as stock prices, valuation metrics, etc.) using a third-party API. This ensures that the data used is always up to date.

3. You enter a stock ticker
   The user types in a stock symbol (like AAPL or TSLA) and starts the process.

4. It sends the ticker to the spreadsheet
   The tool updates the spreadsheet with the ticker you entered, which trigger formulas already set up in the sheet.

5. It performs calculations
   The script runs valuation methods like discounted cash flow (DCF) and relative (peer-based) valuation. It also filters out invalid values.

6. It fetches valuation results
   After processing, the tool reads key valuation outputs from the spreadsheet, such as price estimates, growth rates, and comparable company metrics.

What It Includes

- Peer comparison: In addition to discounted cash flow (DCF) analysis, the sheet includes relative valuation metrics by comparing the company to similar peers (based on EV/EBITDA, ROIC, etc.).

- Averaged growth rates: The DCF estimate is based on the average of three different growth rate estimates: historical average, compound annual growth rate (CAGR), and a linear forecast. This helps make the estimate more stable and consistent.

What You See

- A field to enter the stock ticker
- A button to start the valuation
- A summary of the valuation results, such as:
  - Estimated price
  - Growth expectations
  - Valuation based on multiples and peer comparison

Example Use

You open the app, enter a stock symbol, click the button, and get a quick summary of the stock's estimated value based on a live-updated spreadsheet model that combines company fundamentals with peer-based metrics.
