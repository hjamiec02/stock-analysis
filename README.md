# Refactored Stock Analysis with VBA
## Overview of Project
  
Purpose
  
  The purpose of this project was to refactor code in VBA, originally used to analyse specific green stocks for the years of 2017 and 2018.  The code was refactored in an effort to be more efficient.

## Results

Comparisons

  I was provided with a spreadsheet containing the ticker name, starting date, opening, high, low, closing and adjusted closing price, as well as the volume of each stock. During the module I worked my way through creating a macro that provided each ticker, total daily volume and the percentage return on each stock. The analysis of the green stocks from 2017 to 2018 showed stocks in this category perfored significantly better in 2017 compared to 2018, as seen below: <img width="842" alt="Screen Shot 2022-05-21 at 4 24 42 PM" src="https://user-images.githubusercontent.com/105119531/169668381-a673286f-c1c6-44e6-8b95-7f28abaf0ce4.png">
<img width="791" alt="Screen Shot 2022-05-21 at 4 25 14 PM" src="https://user-images.githubusercontent.com/105119531/169668385-a16acb12-ebfc-48f8-8788-47099afaf031.png">

Coding

  I was provided with an outline, used the code from the module and refactored it to be more effecient in providing me with the same analysis. The refactored code was was faster in both cases. 
    
    Time to run code 2017 original:   0.3515...
    Time to run code 2017 refactored: 0.0781...
    Time to run code 2018 original:   0.2812...
    Time to run code 2018 refactored: 0.0859...
  
Code included below:
  
   '1a) Create a ticker Index

    tickerIndex = 0

    '1b) Create three output arrays
    
    Dim tickerVolumes(12) As Long

    Dim tickerStartingPrices(12) As Single

    Dim tickerEndingPrices(12) As Single

    ''2a) Create a for loop to initialize the tickerVolumes to zero.

    For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
     
        'increase tickerVolumes
    
    Next i

     ''2b) Loop over all the rows in the spreadsheet.
    
    For i = 2 To RowCount

        '3a) Increase volume for current ticker
            
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex
        'If  Then
            
        If Cells(i - 1, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
              
              tickerStartingPrices(tickerIndex) = Cells(i, 6).Value

        End If
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
              
              tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
              
        End If
        
            '3d Increase the tickerIndex.
        
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
            
            End If

        'End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
    
 ## Summary
 
 
