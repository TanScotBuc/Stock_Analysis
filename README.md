# Stock Analysis

## Overview

### Purpose
My friend Steve needed my help to analyze stock market data in order to help advise his parents, who are preparing to invest. I already had a usable code for analyzing a few companies. Using that as a base I refactored the code into something that could analyze thousands of companies in one loop. The refactored code was far more efficent and even integrated some formatting to make the final product more readable.
### Results
After running my analysis it was very easy to see that the stock market overall performed much worse in 2018 than in 2017. The data for 2017 ,linked [here](Stock_Analysis_Visual_Aids/all_stocks_2017.png), shows that only one company had a negative return on the year. Meanwhile, the data for 2018, linked [here](Stock_Analysis_Visual_Aids/all_stocks_2018.png) paints a very different picture, only showing all but two companies with a negative return on the year.
## Summary

### Pros and Cons
Refactoring the original code had a significant effect on the speed at which the subroutine completed. The timing reports before refactoring can be viewed in these links. ( [2017](Stock_Analysis_Visual_Aids/nonrefactored2017.png) [2018](Stock_Analysis_Visual_Aids/nonrefactored2018.png) )
The refactored code works more than six times faster than the original due to the addition of this piece of code that completed all my analysis in one loop rather than running one loop for each company. 
```
    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
        
    If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j - 1, 1).Value <> tickers(tickerIndex) Then
      tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
    End If
            
    If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
      tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
      tickerIndex = tickerIndex + 1
    End If
```    
The refactored code is without a doubt more efficient and flexible. However for analyzing the data set that I was given, I believe the original code would serve my purposes just fine.
