# Stock Analysis Challenge

## Overview of Project

In this project, I aim to conduct a stock market performace analyses through scripting visual basic for application (VBA) codes using Microsoft Excel. To accomplish the analyses, a dataset from WallStreet is used in which stock-specific information such as ticker name, stock value (openning, closing, and highest values for each date), and daily trade volume of different stocks are stored. The results will help ivestors to identify stocks with promising performance and decide on available investment options.

### Purpose

To help Steve's parents to decide on available which investment options.

## Analysis

### VBA Code

Following represents the refactored script for the stock performance analysis:

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


            Worksheets(yearValue).Activate
                
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

In general, refactoring code allows for restructuring an existing code without changing its external behavior. This process is normally expected to be associated with severla benfits such as improving code performance, making code easier to read, and facilitating onboarding New developers. In fact a refactored code is a simplified and clearer version of the original code which help developers to spend less on debugging or understanding the code and more on other tasks such as developing new features. The application run by a refactored code could also, due to better performance, enhance usability and reliability from the user perspective. Notwithstanding all the advantages counted for refactoring code, this strategy may not be useful in all circumstances. Specifically, code refactoring is a time-consuming process and thus in cases such as facing an approching deadline it may be wiser if developers spend their time on more crucial tasks. The time-consuming nature of code refactoring also makes refactoring lecagy codes very expensive and an infeasible strategy. Overall, refactoring code should be of priority when chances of enhancements are high, code smell is detected (e.g., long method, duplicate code), or large number of bugs are raised due to low quality codes.

### How do these pros and cons apply to refactoring the original VBA script?

The major influence of refactoring on the original VBA code written throughout this module was improving the performance of "For" loops by defining an indexing variable which allows to loop through all the data one time. In fact in the initial version, the VBA code would conduct analysis for each ticker separately and run it 12 times. However, the refactored code would do all the calculations for all the tickers at once and store them together. This way, the code spend 11 times less by looping the entire rows once. Although this structural change looks reasonable and expected to improve performance, the low overall running time for both coding structures suggests that, following the logic discussed in the previous question, this single refactoring action may not justify its essentiality in case there were other tasks at hand with high priority.

