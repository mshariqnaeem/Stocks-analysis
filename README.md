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
**8.** Use a for loop through the arrays created to output the "Ticker", "Total Daily Volume", and "Return" Columns in the Spreadsheet

## Results

### Original Results

The original results and the time it took to reach the results using the original VBA code is displayed below

![Old_2017_Analysis](https://user-images.githubusercontent.com/92459399/141696705-37ad70aa-af0a-4ea2-af23-172964c583a4.PNG)

![Old_2018_Analysis](https://user-images.githubusercontent.com/92459399/141696712-2165b92b-a4c8-4f9f-aae9-ece98796b660.PNG)

We can see that in both cases the runtime of the analysis is roughly around 1.30 seconds or more. Although this is not a long time by any means we wanted to see if it was possible to complete the same analysis in a more time efficient manner by refactoring the code.

Using the rules discussed in the background section I started to refactor the code in the following steps

**Step 1**

As asked in the first step I created a variable named tickerIndex and set the variable to the value zero.

![Step_1](https://user-images.githubusercontent.com/92459399/141697026-2e0ecc57-b2d6-4101-8605-e7f6d2b83d5a.PNG)


**Step 2**

As per step 2 we created 3 output arrays named "tickerVolumes", "tickerStartingPrices", and "tickerEndingPrices". we declared the arrays as Long and single to declare the type of value assigned for each array. Note the number "(11)" at the end of each array. This number denotes the number of values each array can store with the number 11 able to store values from 0 to 11. As we have 12 total tickers this is why we include the 11 in our parentheses.

![Step_2](https://user-images.githubusercontent.com/92459399/141697200-9f9067a9-fb6f-428a-b095-d677022dbf81.PNG)

**Step 3**

In this step we create a for loop to initialize the tickerVolumes variable to zero. The For loop will run for iterations from 0 to 11 for the 12 different tickers we have included in our code from zero to 11.

![Step_3](https://user-images.githubusercontent.com/92459399/141697450-39d08af4-b816-4d55-baf5-066e798d4eaa.PNG)

**Step 4**

In this step we loop over all the rows in the spreadsheet and increase the volume for the current ticker using an If-Then statement. We also declared tickerIndex as a value in this if statement to help us accomplish this goal

![Step_4](https://user-images.githubusercontent.com/92459399/141697536-20935e2d-4665-4d5d-a2ee-a3d56eb4a5ea.PNG)

**Step 5**

In this step we use an If-Then statement to find out if the current row we are on is the first row. We use an "And" in this statement to help us accomplish this goal using tickerIndex as a value in the statement.

![Step_5](https://user-images.githubusercontent.com/92459399/141697635-b93dd3f6-bed0-4c50-90a7-c899dedbb767.PNG)

**Step 6**

We Essentially repeat the last step and use an If-Then statement to find the last row. Once again we used "And" to help us accomplish this goal.

![Step_6](https://user-images.githubusercontent.com/92459399/141697718-4976f618-669b-4af7-9d20-19c708aadc9c.PNG)

**Step 7**

In this step we check if the next row's ticker doesn't match the previous row's. If this happens then the value of the ticker increases via the tickerIndex value.
Once again we use a If-then stateement along with the "And" function to accomplish this.

![Step_7](https://user-images.githubusercontent.com/92459399/141697870-237e689b-342c-4fe7-9e6a-dfde72f17209.PNG)

**Step 8**

In this final step we juse a For loop to loop through the array's and output the "Ticker", "Total Daily Volume", and "Return" into the "All stocks analysis" worksheet. We do this with the Cells function and pulling the data from the "tickers(i)", "tickerVolumes(i)", "tickerStartingPrices(i)" and "tickerEndingPrices" arrays we created and used in the loops prior.

![Step_8](https://user-images.githubusercontent.com/92459399/141697919-23585191-2b9d-4435-8aa5-a2bf7f6559ad.PNG)

### End Result

Now when we run our refactored Stock Analysis we get the following results when we run the Analysis for 2017 and 2018. As we can see the run time for both years is more than 1 second faster afer refactoring the original VBA code used. The main reason for this is the original code had 1 long for loop and would need to run the code line by line to reach the final result. In our refactored code we are using shorter for loops and processing data more efficiently which results in the quicker run time for the stock analysis.

![VBA_Challenge_2017](https://user-images.githubusercontent.com/92459399/141698287-373c6103-db75-44e4-9f60-3f25647d6283.PNG)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/92459399/141698293-712a0c1a-426e-4f70-91e4-3573190c64d5.PNG)


### Summary

In summary refactoring code can make existing code run much more efficiently if we understand what the existing code is doing and what we need to accomplish. We can complete what multiple subroutines were accomplishing in one subroutine as well. For example in the previous analysis prior to refactoring, we had multiple subroutines created for the analysis and formatting of the results. In our refactored code we can accomplish this in one subroutine.

The advantages of the original factored code was that we had one for loop which accomplished the goal of completing our stock analysis and the disadvantage of the time it would take to complete the analysis. Our refactored code has the advantage that we can run the analysis quicker, however this is at the cost of spending the additional time to refactor our code to run more efficiently. This was accomplished by creating new array's which took a little more time to code and understand, however the end result for the analysis was reached in a much more efficient and timely manner
