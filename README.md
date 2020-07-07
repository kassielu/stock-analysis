# Automating Stocking Analysis with VBA Code

## Overview of Project
In this project we're helping a newly Finance graduate, Steve, with analysing stock data for his first clients, his parents. Steve's provided us with a handful of green energy stocks data in an Excel file. With the data, we have created macros to automate the analasis process for Steve to effectively analyse all kinds of stocks in a very short amount of time.

### Purpose
To automate the stock analysis process with macros using Excel VBA to reduce time and decrease errors.

## Analysis and Challenges
We started our initial analysis with an Excel file called [Green Stocks] (green_stocks.xlsm). Within that file, the tab called DQ Analysis, contains stock analysis for the DQ stock. Although his parents intially wanted to only invest in DQ stocks, our findings for DQ stock was down 63% in 2008. Hence, Steve would like to analyse other green energry stocks, that will provide a better return for his parents. The All Stocks Analysis tab, shows stock performances for all 12 stocks. In addition, to automate the process further we've created input boxes for Steve to analyse the stocks by year and added buttons to further reduce the amount of time for his analysis.

To further improve the performance for Steve, we have creaetd another Excel workbook called [VBA Challenge] (VBA_Challenge.xlsm). This workbook contains refactored VBA code that drastically reduces the performance time for 2017 from .59 seconds to .10 seconds and for 2018 from .60 seconds to .12 seconds. The [VBA Challenge 2017] (VBA_Challenge_2017.png) and [VBA Challenge 2018] (VBA_Challenge_2017.png) illustrate the performance results.

### Advantages and Disadvantages of Refactoring Code
The advantages on using the refactored code is reducing the performance time. Steve can now analyse many stocks in a very short amount of time. Not only have we helped Steve with reducing his time on stock analysis, we were able to help him automate repeative task that can now be executed in seconds. However, this code is still using hard coding to store the 12 green stocks that Steve provided. To further eliminate any hard coding from our code we should find a way to remove stock names. That way, if Steve wants to analyse other stocks he doesn't have to modify the VBA code. In addition, our code only works if the stocks are sorted in order. Once the sort order is not on the Ticker column, our code runs into a problem and will not work.
