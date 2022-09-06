# Stock Analysis

## Overview of Project

### Purpose

The purpose of Refactoring Code is to help Steve expand through the data set in VBA to include the entire stock market of Wall Street over the last couple of years. By refactoring we will if the code will be more efficient.  

## Results 

### Refactored Code

Sub AllStocksAnalysisRefactored()
    
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("AllStocksAnalysis").Activate
    
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
    
    '1a) Create a ticker Index
        tickerIndex = 0

    '1b) Create three output arrays
        Dim tickerVolumes(12) As Long
        Dim tickerStartingPrices(12) As Single
        Dim tickerEndingPrices(12) As Single

    '2a) Create a for loop to initialize the tickerVolumes to zero.
        For i = 0 To 11
            tickerVolumes(i) = 0
            tickerStartingPrices(i) = 0
            tickerEndingPrices(i) = 0
        Next i
        
    '2b) Loop over all the rows in the spreadsheet.
        For i = 2 To RowCount
    
    '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value

    '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If       
           
       'End If
        
     '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
         'If  Then
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
         End If
            
     '3d) Increase the tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
        End If
        'End If
    
        Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
        For i = 0 To 11
            Worksheets("AllStocksAnalysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
    
    'Formatting
    Worksheets("AllStocksAnalysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
        Cells(i, 3).Interior.Color = vbGreen Else
        Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
    End Sub

    Sub ClearWorksheet()
    Cells.Clear
    End Sub

### Results from code

In 2017 the results came back as expected, and the code ran in 0.1796875 seconds

![2017 results](https://user-images.githubusercontent.com/111101012/188550877-d7147f51-7416-4591-9080-12944e0715aa.png)

![2017 time](https://user-images.githubusercontent.com/111101012/188550893-b192b69c-653f-4984-90e7-8963bdbf924f.png)

In 2018 the results came back as expected, and the code ran in 0.1875 seconds

![2018 results](https://user-images.githubusercontent.com/111101012/188550905-61132ecb-5324-4cf2-878f-0660b91c7658.png)

![2018 time](https://user-images.githubusercontent.com/111101012/188550917-05feb46a-8445-42ad-a07c-d6d267dd1de2.png)

## Summary

### Advantages and disadvantages of refactoring code in general

Advantages of refactoring code in general are that we can improve, and personalize a code to our needs, however a disadvantage is that when we refactor a code that is not made by us, we need to fully understand each part of the code, and what it does so we can make changes, and come back with the results we need in an efficient way. 

### Advantages and disadvantages of the original and refactored VBA script

Advantages of the original code is that it seemed simpler, however it was not built to analyze a large dataset. The refactored VBA script was faster, however it was more complex to understand. 
