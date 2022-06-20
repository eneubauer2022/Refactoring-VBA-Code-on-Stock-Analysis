# Refactoring VBA Code on Stock Analysis

## Overview of Project
I was asked to refactor a macro that was written in response to analyzing how green stocks performed on the market over several years. This would help Steve (and his parents) decide where they should invest their money in the future. The goal was to to loop through all the data at once, hopefully creating a more efficient macro that produced results in a shorter time span than the original code. By doing this, Steve could analyze all 11 stocks and their data quickly in a clean, clear format to present to his parents. 

## Results
From the results, we are able to see that the stock performed better in 2017 compared to 2018. In 2017, all but 1 stock (TERP) had a positive return. However in 2018, all but two stocks (ENPH & RUN) had negative returns. The refactored code did end up producing faster results. The results from 2017 ran 81% faster than when performed with the original code, while the refactored code for 2018 ran 84% faster. This will make it easier to analyze all the data and make a better judgement call on which stocks to choose in the future. 

![this is an image](https://github.com/eneubauer2022/Refactoring-VBA-Code-on-Stock-Analysis/commit/41fbcca9fbcb46c5a50c8910d81f4bbcc4128319#diff-e25a5974a639f075d3c5cdb4b96a17f77334d746c57681349e2d91ab0aa93146)

## Code
```
') Create a ticker Index
    tickerIndex = 0
    

    ') Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
        
    
    ') Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(1) = 0
    
    Next i
       
    
        
    ') Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        ') Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        
        ') Check if the current row is the first row with the selected tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
         
        End If
        
        ') check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
         End If
         

            ') Increase the tickerIndex.
          If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
        
        End If
    
    Next i
    
    ') Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
           
        
        
    Next i
```

##
