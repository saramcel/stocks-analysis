# Stocks Analysis Challenge
### An analysis of green stocks to advise Steve's parents on investments, refactored for speed.

# Project Overview

## Purpose

The purpose of this analysis is to quickly present information about the trading volume and the yearly return for 12 green stocks using data from 2017 and 2018.  

## Background 

Steve wants to help his parents invest in green stocks. They started with one stock of interest and they would like to make comparisons to choose a good investment. We are helping create code that will quickly calculate our key indicators for several candidtate green stocks that Steve has picked. The first time we tried to write this code, it was slow. We have refactored it to make it faster.

# Results

## Analysis Results

The green stocks that Steve selected changed performance from 2017 to 2018. 

![VBA stocks analysis 2017 result](https://github.com/saramcel/stocks-analysis/blob/1964cbe9bce9b7ca64f2a125c185ef4f24889571/Resources/Results_2017.png)

Most of the stocks, with the exception of TERP, showed a positive return in 2017. The stocks DQ, ENPH, FSLR, and SEDG all had yearly returns over 100% in 2017. The highest traded stocks in 2017 were SPWR, FSLR, CSIQ, and RUN. 

![VBA stocks analysis 2018 result](https://github.com/saramcel/stocks-analysis/blob/1964cbe9bce9b7ca64f2a125c185ef4f24889571/Resources/Results_2018.png)

In 2018, only two of the stock choices had positive yearly returns, ENPH and RUN. The highest traded stocks in 2018 were ENPH, SPWR, RUN, and FSLR. 

## VBA Code Comparison

The refactored code ran much faster than the previous code that was developed during the asynchronous modules. While the first code was valuable to teach nested loops, it ran very slowly. The refactored showed how to use arrays to avoid having to loop so many times.

## Original Code

There were nested loops in the original VBA script. The first loop reset variables for each new ticker and printed the results from the nested second loop. The second loop went through every line of the data set to check the if-then conditions. The second loop kept going to the end of the data sheet, even when we had already found the ending price. 

**Loop through all data 12 times**

```
 For i = 0 To 11
    
    ticker = tickers(i)
    totalVolume = 0
    
    For j = 2 To RowCount
        'activate data worksheet
        Worksheets(yearValue).Activate
        
        'increase totalVolume if ticker matches ticker from the array
        If Cells(j, 1).Value = ticker Then
            totalVolume = totalVolume + Cells(j, 8).Value
        End If
        
        'set starting price as first row of ticker
        If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
            startingPrice = Cells(j, 6).Value
        End If
        
        'set ending price as final row of ticker
        If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
            endingPrice = Cells(j, 6).Value
        End If
        
    Next j
    
    'output results
    Worksheets("All Stocks Analysis").Activate
    
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = totalVolume
    Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
    
Next i

```  

It wasted time to keep looping through every row of data even after the end cell had been found and assigned. 

## Refactored Code

The refactored code uses arrays to store the results for each ticker and then prints them all in cells at the end. Using an array is useful because we only need one loop statement that goes through the data. The `tickerIndex` variable keeps the place of all the arrays at once. This variable advances after the end price is found, and because our ticker array and our ticker data are arranged in alphabetical order, this brings up the next ticker. We only have to loop through the data one time rather than 12 times.

**When we have the ending price of the current ticker, start looking for the next ticker**

```
       If tickers(tickerIndex) = Cells(i, 1).Value And tickers(tickerIndex) <> Cells(i + 1, 1).Value Then
             
             'store the last closing price of the year
             tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
             
             '3d Increase the tickerIndex.
             tickerIndex = tickerIndex + 1
        
        End If
```

## Run Time Improvement

The refactored code was much faster than the original code for both years of analysis. The user can perceive the difference in speed. Please see resulting message boxes below for details. 

### 2017 Analysis

**Original Time**

![Original VBA stocks analysis 2017 run time](https://github.com/saramcel/stocks-analysis/blob/1964cbe9bce9b7ca64f2a125c185ef4f24889571/Resources/VBA_Original_2017.png)

**Refactored Time**

![Refactored VBA stocks analysis 2017 run time](https://github.com/saramcel/stocks-analysis/blob/1964cbe9bce9b7ca64f2a125c185ef4f24889571/Resources/VBA_Challenge_2017.png)

### 2018 Analysis

**Original Time**

![Original VBA stocks analysis 2018 run time](https://github.com/saramcel/stocks-analysis/blob/1964cbe9bce9b7ca64f2a125c185ef4f24889571/Resources/VBA_Original_2018.png)

**Refactored Time**

![Refactored VBA stocks analysis 2018 run time](https://github.com/saramcel/stocks-analysis/blob/1964cbe9bce9b7ca64f2a125c185ef4f24889571/Resources/VBA_Challenge_2018.png)


# Summary

### What are the advantages or disadvantages of refactoring code?

Advantages of refactoring code are that a new set of ideas can make the code run more quickly and efficiently. A slight change to the design pattern and can create code that scales more easily to larger data sets. One disadvantage of refactoring code is that some of the code might break with no apparent fix, especially original code that is smelly and really has no reason to work but somehow does. Another disadvantage is that there might be some assumptions about the data that are not made explicitly clear in the original code, like how it is sorted, that would prevent new code from working. 

### How do these pros and cons apply to refactoring the original VBA script?

The refactored script worked much faster than the original script because it was more efficient and ran through the datasheet loop only once. This script would be faster to use with a large dataset if two conditions are met: The data is sorted ascending by date, and the data is sorted ascending by ticker. The disadvantage is that the refactored code will not work if the data is not sorted properly, which was also an issue with the original VBA script. However, there is probably a way to sort the data using VBA. 
