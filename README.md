# Stocks-Analysis

## Overview of Project
### Background
#The purpose of this project was to use Visual Basic for Applications (VBA) in Excel to help out my friend Steve who had just graduated with a 
#finance degree. His parents decided to be his first clients and wanted to invest in some sort of green energy. They hadn't done too much 
#research, so Steve's Parents decided to invest their money into DAQO New Energy Corp (DQ) who basically make silicon wafers for solar pannels.

### Purpose
#Steve made a promise to look into DAQO, but felt that they needed to diversify their funds. So he wanted to analyze different green energy 
#stocks, as well as DAQO stock. Steve gave me an excel file that had all of the data he needed me to analyze. Using VBA, I created code to 
#help Steve analyze any stock to minimize error and to visually undertand the trends.

## Analysis: 
### Process:
#In order to help steve analyze stocks, I created code in VBA. Steve asked that I find the total daily volume (number of shares traded throughout
#the day) and yearly return for each stock (percentage difference in price from beginning of the year to the end). So I created 3 sections: 
#"Year", "Total Daily Volume" and "Return".

#To understand how actively any stock was traded in a certain year, I summed up all of the daily volume for DQ to get the yearly volume and a 
#rough idea of how often it was traded. And to understand how a stock performed in a certain year, I calculated the yearly return for any stock by 
#looping through all the rows, checked if the current row was the first row of DQ's data and then, if so, set the starting price to the 
#closing price in the current row. After, I checked if the ticker in the current row was that stock and if the ticker in the previous row was not 
#that stock.

#To run analyses on all of the stocks for any year, not just DQ, I created a program flow that looped through all of the tickers. Also, since 
#Steve may want to look at a different set of stocks in the future, I created a flexible macro for running multiple stocks for any year.

#In the future, Steve may want to perform his analysis on larger datasets so Steve wanted to see how fast his VBA code will compile the results. 
#To help Steve, I added a script that calculated how long the code took to execute.

#I also refactored the code that collected all of the same information in order to see if I could successfully make the VBA script run faster.
#Seen below is the Original Script and the Refractored Script.

### Original Script:
Sub yearValueAnalysis()
    
    Dim startTime As Single
    Dim endTime As Single
    
    'User input
    yearValue = InputBox("What year would you like to run the analysis on?")
    
        startTime = Timer
        
    '1. Format the output sheet on the "All Stocks Analysis" worksheet.
    
    Worksheets("All Stocks Analysis").Activate
    
    'Title
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Headers
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
    '2. Initialize an array of all tickers.
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
    
    '3. Prepare for the analysis of tickers.
        'a.) Initialize variables for the starting price and ending price.
        
        Dim startingPrice As Single
        Dim endingPrice As Single
        
        'b.) Activate the data worksheet.
        
        Worksheets(yearValue).Activate
        
        'c.) Find the number of rows to loop over.
        
        RowCount = Cells(Rows.Count, "A").End(xlUp).Row

    '4. Loop through the tickers.
    For i = 0 To 11
        
        ticker = tickers(i)
        totalVolume = 0
        
    '5.Loop through rows in the data.
    
        Worksheets(yearValue).Activate
        For j = 2 To RowCount
        
        'a.) Find the total volume for the current ticker.
        
            If Cells(j, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value
            End If
        
        'b.) Find the starting price for the current ticker.
        
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                startingPrice = Cells(j, 6).Value
            End If
        
        'c.) Find the ending price for the current ticker.
        
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                endingPrice = Cells(j, 6).Value
            End If
        
        Next j
        
    '6. Output the data for the current ticker.
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
    Next i
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.Bold = True
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("A3:C3").Font.ColorIndex = 3
    Range("A3:C3").Font.FontStyle = "Underlined"
    With Worksheets("All Stocks Analysis").Range("A3:C3")
        .Font.Size = 14
    End With
    
    
    Range("B4:B15").NumberFormat = "$#,###.00"
    Range("C4:C15").NumberFormat = "0.00%"
    
    Columns("B").AutoFit
    
    'Color coding the data
    dataRowStart = 4
    dataRowEnd = 15
    
    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
    
            'Color the cell green
            Cells(i, 3).Interior.Color = vbGreen
    
        ElseIf Cells(i, 3) < 0 Then

            'Color the cell red
            Cells(i, 3).Interior.Color = vbRed
    
        Else
            'Clear the cell color
            Cells(i, 3).Interior.Color = xlNone

        End If
    Next i
    
        endTime = Timer
        MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
        
End Sub

### Refractored Script:
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
    For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
        Worksheets(yearValue).Activate
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
            
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        End If
            
            '3d) Increase the tickerIndex.
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

### Results
#Comparing the time for both the original and the refactored script, it is very clear that the refractored script runs faster.

### Original Script Times:

<img width="408" alt="PREVBA_Challenge_2017" src="https://user-images.githubusercontent.com/85847344/124397472-d75f9b80-dcc4-11eb-8714-dd2112870eca.png">
<img width="400" alt="PREVBA_Challenge_2018" src="https://user-images.githubusercontent.com/85847344/124397473-d7f83200-dcc4-11eb-950e-bae4e45b7c4d.png">

### Refractored Script Times
<img width="409" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/85847344/124397479-e0506d00-dcc4-11eb-8e2a-afbd5971f83f.png">
<img width="400" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/85847344/124397482-e21a3080-dcc4-11eb-9a14-f5cdd866245b.png">

#Now, looking at the actual results of the stock performance between 2017 and 2018, it is very interesting.


Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script 
and the refactored script.
The analysis is well described with screenshots and code




## Summary
#Summary: In a summary statement, address the following questions.

advantages and disadvantages of refactoring code in general
#Advantages: more consise. runs faster
#Disadvantages: if too refractored it can be hard to debug because of not knowing which part is going wrong. harder to look at code for fist time and understand what is going on
helps keep it manageable without major overhauls but may not set the app up for new development technologies or application languages. Rewriting code enables foundational changes to the code but risks confusing developers or even breaking the product.

advantages and disadvantages of refactoring this stock code
#Advantages: The largest advantage is the time
#Disadvantages: Steve may not be able to understand how the code is working as well, but that shouldn't really matter since the coding part isn't #his issue

