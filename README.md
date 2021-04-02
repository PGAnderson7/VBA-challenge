# VBA Challenge Code Description
This code runs in a file full of yearly stock data by day that needs to be sorted alphabetically by ticker name prior to running.  For each worksheet it summarizes year change, percent change and total stock volume per ticker name.  As well as tracking greatest % increase, greatest % decrease, and greatest total volume.

This is accomplished by stepping through each worksheet row by row and tracking the open price, close price, and volume values per day.  When it reaches the end of each ticker name, it takes the values tracked, calculates the the summarized data, and prints to a summary table within the current worksheet.

Calculations used:
-Yearly Change = Latest close price - earliest open price.
-Percent change =
-Total Stock Volume