VBA: Stock Data
-------------------------------------------------------------------------------------------------
## The data
-------------------------------------------------------------------------------------------------
    The Excel workbook has 4 worksheets.  Each worksheet contains stock data per quarter of 2022 
    (i.e. worksheet 1 corresponds to quarter 1, worksheet 2 to quarter 2, worksheet 3 to quarter 3, 
    and worksheet 4 to quarter 4). 

The stock data includes multiple tickers' open price, close prices, high price, low price 
and stock volume for multiple consecutive dates throughout the quarter.  

-------------------------------------------------------------------------------------------------
## The code 
-------------------------------------------------------------------------------------------------
    For each ticker, the code extracts the open price at the beginning of the quarter 
    and close price at the end of the quarter.  These are then used to calculate the amount
    the price has changed and percent the price has changed from open to close of the quarter.  

    Meanwhile, the code also calculates the total stock volume per ticker per quarter

    Next, the code finds the tickers which had the greatest percent increase, 
    greatest percent decrease, and the greatest total volume per quarter.  

    Finally, the code formats the quarterly change data so that the tickers which the 
    price increased from open to close are green and the ones that decreased are red.  

    The percentages in the data were formatted as percents rounded to the second 
    decimal place.

