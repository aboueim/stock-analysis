# Stock Analysis Challenge

## Overview of Project

In this project, I aim to conduct a stock market performace analyses through scripting visual basic for application (VBA) codes using Microsoft Excel. To accomplish the analyses, a dataset from WallStreet is used in which stock-specific information such as ticker name, stock value (openning, closing, and highest values for each date), and daily trade volume of different stocks are stored. The results will help ivestors to identify stocks with promising performance and decide on available investment options.

### Purpose

To help Steve's parents to decide on which tickers to invest.

## Analysis

### VBA Code

Following represents the final script for the stock performance analysis:

Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
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
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    'Set the starting row number
    startingRow = 2
    
    
    '1a) Create and initialize a ticker Index
    
    Dim tickerIndex As Double
    
    tickerIndex = 0
    

    '1b) Create three output arrays
    
    
    Dim tickerVolumes(12) As Long
    
    Dim tickerStartingPrices(12) As Single
    
    Dim tickerEndingPrices(12) As Single
    
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    
    
    For i = 0 To 11
    
        ticker = tickers(tickerIndex)
        
        tickerVolumes(tickerIndex) = 0
    
            
    ''2b) Loop over all the rows in the spreadsheet.

    
            For j = startingRow To RowCount
    
    
        '3a) Increase volume for current ticker
        
                If Cells(j, 1) = ticker Then
                
                    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
        
                End If
                
            
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            
            
                If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
                
                    tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
            
            
            
        'End If
                
                End If
            
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            
                
                If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
                
                    tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
                
                    
                End If
            
            

            '3d Increase the tickerIndex.
            
                If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
                
                tickerIndex = tickerIndex + 1
            
            
        'End If
        
                End If
        
        
            Next j
        
        
    
    Next i
    
    
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    
    For i = 0 To 11
        
        
        Worksheets("All Stocks Analysis").Activate
        
    
        'output data for Ticker, Total Daily Volume, and Return
        
        Cells(i + 4, 1).Value = tickers(i)
        
        Cells(i + 4, 2).Value = tickerVolumes(i)
        
        Cells(i + 4, 3).Value = (tickerEndingPrices(i) / tickerStartingPrices(i)) - 1
        
        
        
    Next i
    
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
    

End Sub

Sub clearworksheet()


    'clear worksheet
    
    Worksheets("All Stocks Analysis").Activate
    
    
        Cells.Clear
    
        
    
End Sub


### Stock Performance 2017

![This is an image](/VBA_Challenge_2017.png)



### Analysis of Outcomes Based on Goals

![This is an image](/VBA_Challenge_2018.png)


## Summary

### What are the advantages or disadvantages of refactoring code?


### How do these pros and cons apply to refactoring the original VBA script?







