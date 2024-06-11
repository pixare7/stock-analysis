
# Stock Data Analysis VBA Script

## Overview

This VBA script processes stock data across multiple worksheets in an Excel workbook. It calculates the quarterly change, percent change, and total stock volume for each ticker symbol. Additionally, it identifies the greatest percent increase, greatest percent decrease, and greatest total stock volume for the tickers and highlights these values with conditional formatting.

## Setup and Usage

1. **Open the VBA Editor:**
   - Press `Alt + F11` to open the VBA editor in Excel.

2. **Insert the Script:**
   - Copy the entire VBA script and paste it into a new module.
     - In the VBA editor, go to `Insert > Module` to create a new module.
     - Paste the script into the module.

3. **Run the Script:**
   - Press `F5` or go to `Run > Run Sub/UserForm` to execute the script.

## Script Explanation

### Variables and Headers

- **Variables:** The script initializes several variables to store data such as the ticker symbol, open price, close price, quarterly change, percent change, and total stock volume.
- **Headers:** It sets up headers in columns I to L for the output data: Ticker, Quarterly Change, Percent Change, and Total Stock Volume.

### Processing Each Worksheet

The script loops through each worksheet in the workbook and processes the stock data:

1. **Loop through Rows:**
   - It iterates through each row in column A to identify different tickers.
   - When a new ticker is found, it calculates and records the quarterly change, percent change, and total stock volume for the previous ticker.
   - Resets and updates variables for the next ticker.

2. **Greatest Values:**
   - It identifies the greatest percent increase, greatest percent decrease, and greatest total stock volume among all tickers.

### Conditional Formatting

- **Quarterly Change Formatting:** It applies conditional formatting to the quarterly change values in column J.
  - Negative values are highlighted in red.
  - Positive values are highlighted in green.
- **Percentage Formatting:** Percent changes are formatted to two decimal places in column K.
- **Greatest Values Formatting:** Formats the values for greatest percent increase and decrease, and the total volume in the summary section.

### Summary Section

- The script populates a summary section with the greatest percent increase, greatest percent decrease, and greatest total stock volume along with their respective tickers.

### Sample Output

After running the script, the output will be displayed in the following columns:

| Ticker | Quarterly Change | Percent Change | Total Stock Volume |
|--------|------------------|----------------|--------------------|
| AAPL   | 10               | 5.00%          | 50000              |
| MSFT   | -2               | -1.50%         | 30000              |
| ...    | ...              | ...            | ...                |

| Greatest Percent Increase | Ticker | Value |
|---------------------------|--------|-------|
| 15.00%                    | AAPL   | 0.15  |

| Greatest Percent Decrease | Ticker | Value |
|---------------------------|--------|-------|
| -10.00%                   | MSFT   | -0.10 |

| Greatest Total Volume     | Ticker | Value |
|---------------------------|--------|-------|
| 1000000                   | TSLA   | 1E+06 |

---

This script provides an automated way to analyze stock data and highlight key metrics in an Excel workbook, making it easier to visualize and interpret the data.
