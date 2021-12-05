# Refactoring Stock Analysis
## Overview
  This week we learned how to analyze data using Visual Basic for Applications. The goal of this anlysis was to uncover the best company to invest stock in. To best assist Steve I created a macro in order to reveal which stocks did the best between 2017 and 2018. In addition to creating this Macro i refactored the script to improve the speed in which we get the results of the anlysis. Refactoring makes the code more efficient by looping through the code at a significantly faster rate. 
## Results
  In order to complete this anylsis I had to add and delete a bit of code to transition smoother through the data. I first had  initialize a variable that would pull the desired results from the macro. I used "Ticker Index". I created this to access the selected index across the arrays.  Next I created a list of arrays to store data from the data worksheets. The list includes "tickerVolumes, tickerStartingPrices as well as tickerEndingPrices." This allows the script to loop through our data, pulling out the desired values. 
### The refractored code
    Sub AllStockAnalysisRefractored()
    Dim startTime As Single
    Dim endTime As Single
    
    yearValue = InputBox("What year would you like to run the analysis on?")
    startTime = Timer
    
     'Format the output sheet on the "All Stocks Analysis" worksheet.
       Worksheets("All Stocks Analysis").Activate

    Range("A1").Value = "All Stocks(" + yearValue + ")"
    'create header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Volume"
    Cells(3, 3).Value = "Return"
    
    'Initialize an array of all tickers.
    Dim tickers(12) As String
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"

    'Activate the data worksheet.
    Worksheets(yearValue).Activate

    'Find the number of rows to loop over.
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row

    1a) 'Initialize variables for the starting price and ending price.
    tickerIndex = 0
    1b) Create three Output arrays
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    2a) Create a For loop to initialize the tickersVolume to zero
    For I = 0 To 11

     tickerVolumes(I) = 0
    
    Next I
    '2b) Loop over all the rows in the spreadsheet.
     For I = 2 To RowCount
    '3a) Increase volume for current ticker
        
     tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(I, 8).Value
          
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
    If Cells(I - 1, 1).Value <> tickers(tickerIndex) Then
    tickerStartingPrices(tickerIndex) = Cells(I, 6).Value
    End If
            
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
         If Cells(I + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(I, 6).Value

            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1

        End If
    
    Next I
 
    Worksheets("All Stocks Analysis").Activate
    dataRowStart = 4
    dataRowEnd = 15
    For k = dataRowStart To dataRowEnd

    If Cells(k, 3) > 0 Then
       Cells(k, 3).Interior.Color = vbGreen
    
    ElseIf Cells(k, 3) < 0 Then
       Cells(k, 3).Interior.Color = vbRed

    Else
      Cells(k, 3).Interior.Color = xlNone
    End If
        Next k
        
    endTime = Timer
        MsgBox "This code ran in " & (endTime - startTime) & "seconds for the year " & (yearValue)
         'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
        
    End Sub

  After the code successfully looped through the worksheet, I was able to compare the result output times. 
The origional code for 2018 ran at 0.47 seconds. Our refractored code ran at 0.27 seconds. Thats nearly half of the time it took the first go around!
[VBA_Challenge_2018.png.zip](https://github.com/nalicia/stock_analysis/files/7654996/VBA_Challenge_2018.png.zip)
The code for 2017 was similar, Going from running at 0.49 seconds to running at 0.28 seconds! the power of refractoring data produces a m,uch fster result. 
Coincidentaly through this analysis we found that DQ would not be a viable option for steves parents as they dropped 62% from 2017 to 2018. the best option for Steves' parents would be either ENPH returing 81.9% or RUN returning 84% from 2017 to 2018. 
[Refractored Stock Analysis 2017.zip](https://github.com/nalicia/stock_analysis/files/7654998/Refractored.Stock.Analysis.2017.zip)
[Refractored Stock Analysis 2018.zip](https://github.com/nalicia/stock_analysis/files/7654999/Refractored.Stock.Analysis.2018.zip)
debugging cleaner transitioning code, cons are it can be trickjy using vba (Errirs


