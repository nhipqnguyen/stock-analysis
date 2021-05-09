# Stock Analysis with VBA

## Overview of Project

### Purpose
This project is to help Steve, a recent Finance graduate, to analyze stocks in the Green Energy Indsutry in order to make good decisions on stock investments.
Steve's dataset contains data about 12 Green Energy companies. After helping him gain some insights into this specific dataset, we now refactor our macro to make it work for the entire stock market.

## Results

### Comparison between Stock Performance in 2017 and 2018
* Below is the lists of 12 companies in the Green Energy industry and their stock total daily volumes and return in 2017 and 2018

!["Stock Volumes & Returns 2017" Line Chart](https://github.com/nhipqnguyen/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

!["Stock Volumes & Returns 2018" Line Chart](https://github.com/nhipqnguyen/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

* If we look at the "Return" column of both tables, most of the cells on the 2018 table are red while most of the cells on the 2017 are green. This means most of these companies' stock investors earned profits in 2017 but suffered losses in 2018.
* The only 2 companies in the list that returned profits to their investors in both years were ENPH and RUN. It looks like investors are getting more interested in these 2 companies over the years. ENPH had its total daily trading volume increased approximately 2.7 times in 2018 compared to 2017. RUN's total daily trading volume also went up 1.8 times from 2017 to 2018. Therefore, Steve's parents might want to consider shifting their investments from DQ to ENPH and RUN.

### Comparison between Execution Times of The Original Script and The Refactored Script
* The below pop-up messages show the execution times of the original script before refactorization.

!["Execution Time of Original Script" Line Chart](https://github.com/nhipqnguyen/stock-analysis/blob/main/Resources/VBA_Challenge_2017_original_script.png)

!["Execution Time of Original Script" Line Chart](https://github.com/nhipqnguyen/stock-analysis/blob/main/Resources/VBA_Challenge_2018_original_script.png)

* Based on these numbers, the macro ran 4.6 times faster for 2017 data and 5.6 times faster for 2018 data after being refactored. The refactored script is clearly more efficent than the original one.

## Summary

- What are the advantages or disadvantages of refeactoring code?
  *  Refactoring code involves reducing scope, making complex instructions simpler, and combining multiple statements into fewer statements. By cleaning and transforming code, refactoring can make it easier to understand, debug or change software. It also can improve the design of program and its productivity, making executing software faster and more efficiently.
  *  Beside its benefits, code refactoring also has drawbacks. Restructing and cleaning code could be time-consuming. There are also chances that the refactored code causes new bugs and problems.

- How do these pros and cons apply to refactoring the original VBA script?
  *  In the original VBA script, we loop through rows in the data, calculate the total daily volume, starting price and ending price of the current ticker, store those values in 3 variables, print out the results and then reset those 3 variables for the next ticker. The macro has to repeat these steps for many times, which could make it less efficient.

```
For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0
   
        '5) loop through rows in the data
        Worksheets(yearValue).Activate
        For j = 2 To RowCount
        
            '5a) Find total volume for current ticker
            If Cells(j, 1).Value = ticker Then

                totalVolume = totalVolume + Cells(j, 8).Value

            End If
       
            '5b) Find starting price for current ticker
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

                startingPrice = Cells(j, 6).Value

            End If
       
            '5c) Find ending price for current ticker
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

                endingPrice = Cells(j, 6).Value
            
            End If

        Next j
   
        '6) Output data for current ticker
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

    Next i
```
  *  Therefore, we refactored the script by making the macro do fewer steps. Instead of using the same 3 variables and having to reset them for each ticker. We create 3 arrays to hold values for each tickers, which gives each ticker its own storage space for all vaules. This helps the program run faster because we can just loop through the rows, calculate and store the needed values into its own space using an index variable, then print 3 arrays all at once after the loop ends.

```
 For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        If Cells(i, 1).Value = tickers(tickerIndex) Then
            
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        End If
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        
        End If
                         
        '3c) check if the current row is the last row with the selected ticker
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
                
        '3d) If the next row’s ticker doesn’t match, increase the tickerIndex.
            tickerIndex = tickerIndex + 1
        
        End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
```
