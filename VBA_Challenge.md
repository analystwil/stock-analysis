# VBA Challenge

## Overview of Project

The purpose of this analysis was to refactor the code previously used in the module. I wanted to see if refactoring the code would make the script run faster and if it would be more efficient. I also included the entire stock market when completing the refactor instead of the dozen stocks which were used before.

## Results

 The stock performance was better in 2017 with almost all stocks being in the positive. The only stock that was negative in 2017 was TERP. While TERP continued being in the negative in 2018 it increased from -7.2% in 2017 to -5.0% in 2018. In 2018 10 out of the 12 stocks dipped with DQ dipping the lowest with a return of -62.6%. The biggest increase for 2018 was RUN with a return of 84% which was a huge increase compared to the 5.5% return they had in 2017. 

### Refactored Script along with code
 
<img width="368" alt="Screen Shot 2021-03-01 at 2 07 06 PM" src="https://user-images.githubusercontent.com/77358388/109583771-6f0e6c80-7ace-11eb-85d5-a59d723604f2.png">

<img width="302" alt="Screen Shot 2021-03-01 at 2 06 37 PM" src="https://user-images.githubusercontent.com/77358388/109583778-7170c680-7ace-11eb-8b82-bc4d4095de35.png">


 
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
    
    '1a) Create a ticker Index
    
    tickerIndex = 0
   

    '1b) Create three output arrays
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single

    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    
     Worksheets(yearValue).Activate
    
    For i = 0 To 11
    tickerVolumes(i) = 0
    Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
    
    For j = 2 To RowCount
    
        '3a) Increase volume for current ticker
         tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            
            If Cells(j - 1, 1).Value <> tickers(tickerIndex) Then

        tickerStartingPrices(tickerIndex) = Cells(j, 6).Value

     End If
    
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            
            If Cells(j + 1, 1).Value <> tickers(tickerIndex) Then

        tickerEndingPrices(tickerIndex) = Cells(j, 6).Value


            '3d Increase the tickerIndex.
            
            tickerIndex = tickerIndex + 1
        
        End If
    
    Next j
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
    'tickerIndex = ticker(i)
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = (tickerEndingPrices(i) / tickerStartingPrices(i)) - 1

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


### Original Script along with code
 
 <img width="323" alt="Screen Shot 2021-03-01 at 1 49 30 PM" src="https://user-images.githubusercontent.com/77358388/109583739-628a1400-7ace-11eb-9340-2b6980d79017.png">

<img width="299" alt="Screen Shot 2021-03-01 at 1 48 53 PM" src="https://user-images.githubusercontent.com/77358388/109583746-64ec6e00-7ace-11eb-8a6e-3d8bcdc8a9b9.png">


 Sub AllStocksAnalysis()
 
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

       startTime = Timer
    
' 1) Format the output sheet on the "All Stocks Analysis" worksheet.
    Worksheets("All Stocks Analysis").Activate
 
 Range("A1").Value = "All Stocks (2018)"
        
        'Create a header row
        Cells(3, 1).Value = "Ticker"
        Cells(3, 2).Value = "Total Daily Voulme"
        Cells(3, 3).Value = "Return"
        
'2) Initialize an array of all tickers.
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
                  
' 3a) Initialize variables for the starting price and ending price.

    Dim startingPrice As Single
    Dim endingPrice As Single

'3b) Activate the data worksheet.
    
    Worksheets("2018").Activate
    
'3c) Find the number of rows to loop over.

    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
'4) Loop through the tickers.

For i = 0 To 11

    ticker = tickers(i)
    totalVolume = 0

'5) Loop through rows in the data.
    Worksheets("2018").Activate
    For j = 2 To RowCount

  '5a)  Find the total volume for the current ticker.
  
   If Cells(j, 1).Value = ticker Then

        totalVolume = totalVolume + Cells(j, 8).Value

     End If
     
     '5b) Find the starting price for the current ticker.
  
    If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

        startingPrice = Cells(j, 6).Value

     End If
  
  '5c) Find the ending price for the current ticker.
   If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

        endingPrice = Cells(j, 6).Value
        
        End If
    
    Next j
    
'6) Output the data for the current ticker.

    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = totalVolume
    Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

    Next i

        endTime = Timer
        MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub

 The execution times were a lot faster with the refactored script for both 2017 and 2018. 


# Summary

## What are the advantages or disadvantages of refactoring code?

The advantages to me was that you didn't have to recreate the entire code all over again. You could pick and choose what bits and pieces could be used for the refactored code. It tended to be a lot easier for me to navigate with me reusing the same code. The disadvantages would be that I needed to make sure that I went back and changed the correct letters used in my for loops. I had a few mismatches and it took me some time to figure out which lines were incorrect. I also seemed to spend a lot refactoring the code to make sure every line that needed updating was done correctly. 

<img width="541" alt="Screen Shot 2021-03-01 at 2 12 29 PM" src="https://user-images.githubusercontent.com/77358388/109583840-88afb400-7ace-11eb-88f0-6c58c3ed70f4.png">

<img width="492" alt="Screen Shot 2021-03-01 at 2 14 34 PM" src="https://user-images.githubusercontent.com/77358388/109583851-8f3e2b80-7ace-11eb-8e63-39999f2301cc.png">


## How do these pros and cons apply to refactoring the original VBA script?

It was a lot easier to maintain the refactored VBA script because it was my second attempt so I was more familiar with the code. 
