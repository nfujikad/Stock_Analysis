# Stock Analysis with Excel

Click here to view the Excel file: [VBA Challenge - Stock Analysis](https://github.com/nfujikad/Stock_Analysis/blob/main/VBA_Challenge.xlsm)

## Overview of Project 

### Purpose
-   This challenge was to refactor the VBA code for Microsoft Excel and utilize this to run over stocks from 2017 and 2018. In doing so we could determine whether any of these stocks are worth investing in. The refactoring was done to create a fluid and more efficient process.

## Results

### Analysis
-   The analysis was conducted on 12 different stocks in 2017 and 2018 for comparison. The overall goal was to separate and sum the tickers (stocks), the total daily volume, and the return for each stock respective to its year. 
The following is the code displays how variables and array were created and for loops used to separate/sum up the individual tickers and corresponding data.

    '1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    ' If the next row’s ticker doesn’t match, increase the tickerIndex.
    For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    Next i
   
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        '3c) check if the current row is the last row with the selected ticker
        'If  Then
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
         End If

            '3d Increase the tickerIndex.
             If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
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

## Summary

### What are advantages and disadvantages of refactoring code?

-   The refactoring allowed run time to decrease significantly. Typically, what would originally take at least a second now only takes a fraction of the time. Below is shown the run time for the 2018 refactored analysis.

![VBA 2017 Screenshot]( https://github.com/nfujikad/Stock_Analysis/blob/main/Resources/VBA_Challenge_2017.png)

### How do these pros and cons apply to refactoring the original VBA script?
-   Refactoring allows us to clean our data and create a template for similar cases rather than being repetitive. In the end the goal is to be efficient with time and by starting with an organized canvas we can easily troubleshoot and improve code as needed. Additionally, this allows for collaborative work to work faster and smoother. On the other hand, this may not be the best route for those who need quick answers. Refactoring takes time and attention to detail which may not be ideal for someone who only needs to run this code once for a quick answer. This is better suited for those who must frequently run the same analysis.
