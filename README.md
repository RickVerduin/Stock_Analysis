# Refactoring VBA code for Analysis of Stocks

## Overview of Project
The focus of this project is to refactor VBA code used to retrieve stock performance data. The code to be refactored was written to retrieve data from two MS Excel sheets containing opening and closing prices and daily trade volumes for a total of 12 stocks for every trading day in 2017 and 2018. The code asks for input on what year the user would like the analysis for, and creates a formatted table with tickers, total traded volume, and annual returns for each of the stocks. Although the original code performs the task as intended, it has a rather slow processing time, which would be problematic when upscaling the project for use on more than 12 stocks. Refactoring this code will overcome this issue.

## Results
### Changes in code
The original code would loop through the rows in the sheet, use a nested loop to record the values for totalVolume, startingPrice and endingPrice, and subsequently output the values in the table for one ticker at a time. The refactored code uses arrays in combination with the newly created variable tickerIndex rather than individual variables to collect the values for tickerVolumes, tickerStartingPrices and tickerEndingPrices for all tickers, before outputting them in the table.

_Original Code_

![image](https://user-images.githubusercontent.com/93882635/144771828-7771f576-8d4b-4ab4-ad59-7a122d23f135.png)


_Refactored Code_

![image](https://user-images.githubusercontent.com/93882635/144771987-8e3c2d45-daa1-46a9-bae0-0d5a66e1780e.png)


### Changes in Performance
To see if the refactored code is faster than the original code, a timer was added. This timer starts right after the input for the year to be analyzed is received and stops right after the output table is complete. As visible in the following screenshots, the refactored code's execution is significantly faster.

_Original code performance_

![image](https://user-images.githubusercontent.com/93882635/144767915-44878d2b-455e-4280-9c91-00397d833ff3.png)

![image](https://user-images.githubusercontent.com/93882635/144767945-e8d51865-8732-4cbb-8d97-e0bb7dc27e40.png)


_Refactored code performance_

![VBA_Challenge_2017](https://user-images.githubusercontent.com/93882635/144772185-74e5fc3b-5365-4a83-8efc-f1b1fd05dc10.png)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/93882635/144772200-d587e82e-e119-4f01-bbc5-50d8052f4dbe.png)


## Summary
### Advantages and disadvantages of refactoring code
As shown in this project, refactoring the code can lead to more efficient code. In this case it made the code run significantly faster, which makes it a viable option when the amount of stocks that need to be analyzed increases. Refactoring can also simplify code. In this case, the nested loop was removed, which made the code easier to read. Simpler code will be easier to understand and maintain, saving time and money in the future.

The main disadvantage of refactoring code is that it can be time consuming. It requires a thorough understanding of how the original code works, before an attempt at refactoring can be made. Figuring out the exact way the code works and coming up with simplifications can take up large amounts of time. There is also a risk that one breaks the code, rather than improve it. In this case, the code ended up not working on multiple occasions, which required troubleshooting before steps towards the final improved code could be made.
