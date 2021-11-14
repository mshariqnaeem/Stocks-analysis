# Stocks-analysis with VBA Refactored

## Overview of Project
This project takes the data from the worksheets 2017 and 2018 within the green_stocks analysis workbook and outputs the results of the stock tickers by year. This was accomplished using VBA instead of a inputting formula's into the front-end platform of Microsoft Excel.

### Purpose
The purpose of using VBA to accomplish this goal was to firstly understand how to use VBA to analyze not just stock data, but also how to analyze any data set using VBA code. Secondly this particular exercise was to understand how to take an existing VBA code and refactor the code to make it run more efficiently.

**This teaches us how to:**\
**1.** understand how the code works\
**2.** refactor the code to run and accomplish the same goal\
**3.** make the refactored code run more efficiently

### Background
The background to this project is that Steve wanted to analyze stock data in the green_stocks excel workbook for 2017 and 2018. This was to figure out what would be the best strategy for his parents investment portfolio. Originally we took the stocks and created multiple VBA subroutines to display the results on a new worksheet named All Stocks Analysis. The original code however was not very efficient and the purpose of this exercise was to make the existing code run more efficiently with specific sets up rules/parameters which were provided to us.

**The following rules were:**\
**1.** Create a tickerIndex variable and set it to zero\
**2.** Create 3 output arrays names tickerVolumes, tickerStartingPrices, and tickerEndingPrices\
**3.** Create a for loop to initialize the ticker Volumes to zero\
**4.** Inside the for loop creat a script that increases the current tickerVolumes variable and adds the ticker volume to the current stock ticker via the tickerIndex\
**5.** Use an if-then statement to check if the current row is the first row\
**6.** Repeat the the previous step to check if the current row is the last row\
**7.** Write a script that increases the tickerIndex if the next row's ticker doesn't match the previous row's ticker\
**8.** Use a for loop through the arrays created to output the "Ticker", "Total Daily Volume", and "RReturn Columns in the Spreadsheet

## Results

### Original Results

The original results and the time it took to reach the results using the original VBA code is displayed below

![Old_2017_Analysis](https://user-images.githubusercontent.com/92459399/141696705-37ad70aa-af0a-4ea2-af23-172964c583a4.PNG)

![Old_2018_Analysis](https://user-images.githubusercontent.com/92459399/141696712-2165b92b-a4c8-4f9f-aae9-ece98796b660.PNG)

We can see that in both cases the runtime of the analysis is roughly around 1.30 seconds or more. Although this is not a long time by any means we wanted to see if it was possible to complete the same analysis in a more time efficient manner by refactoring the code.

Using the rules discussed in the background section I started to refactor the code in the following steps\
**Step 1**

![Step_1](https://user-images.githubusercontent.com/92459399/141697026-2e0ecc57-b2d6-4101-8605-e7f6d2b83d5a.PNG)
