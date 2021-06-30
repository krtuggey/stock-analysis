# Clean Energy Stocks Analysis

## Overview
Steve has been researching the stocks for clean energy companies in hopes of finding a good investment for his parents. He has enlisted my help to create code that will help him analyze the stocks of 12 companies from the years 2017 and 2018.

## Purpose
The purpose of this analysis, other than to help my friend Steve out, is to find out which clean-energy stock has the best yearly return. My code provides an option for Steve to choose which year (2017 or 2018) he wants to analyze. Then, the code will display the total daily volume and return of each stock of that year. I added conditional formatting provided for better visualization as well. This code works correctly, but I have refactored this code so that it runs faster than the original and handles larger amounts of data in the case that Steve adds more stocks to his datasheet.

## Results
Both  my original code and refactored code included the code to create a timer, create an input box, and format the worksheet. I will be focusing on the remaining code that differs.

##### Original Code
My original code (seen below) uses variables to create a volume count, and a starting and end price. Then, a nested for loop runs through the tickers and rows of data to tally the data needed for the worksheet. 

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
    
    'Prepare for the analysis of tickers
    'Initialize variables for the starting price and ending price.
    Dim startingPrice As Single
    Dim endingPrice As Single
    
    'Activate the data worksheet.
    Worksheets(yearValue).Activate
    
    'Find the number of rows to loop over.
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    'loop through the tickers
    For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0
        
        'loop through rows in the data
        Worksheets(yearValue).Activate
        For j = 2 To RowCount
        
            'find the total volume for the current ticker
            If Cells(j, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value
            End If
            
            'find the starting price for the current ticker
            If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
                startingPrice = Cells(j, 6).Value
            End If
            
            'find the ending price for the current ticker
            If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
                endingPrice = Cells(j, 6).Value
            End If
            
        Next j
        
        'output the data for the current ticker
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
        
    Next i
    
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

##### Refactored Code
My refactored code makes use of one ticker index variable to run through arrays. This process removes the nested for loop. Overall, these changes will speed up the processing time and decrease the memory use.
  
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
    
    'Create a ticker Index
    
    Dim tickerIndex As Single
    tickerIndex = 0
    
    'Create three output arrays
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    'Create a for loop to initialize the tickerVolumes to zero.
    
    For i = 0 To 11
    
        tickerVolumes(i) = 0
        
    Next i
    
    'Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        'Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
            
        'Check if the current row is the first row with the selected tickerIndex.
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
        
        'Check if the current row is the last row with the selected ticker
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
            'Increase the tickerIndex
            tickerIndex = tickerIndex + 1
            
        End If
    
    Next i
        
    'Loop through your arrays to output the ticker, total daily volume, and return
    For i = 0 To 11
    
        Worksheets("All Stocks Analysis").Activate
        tickerIndex = i
        Cells(4 + i, 1).Value = tickers(tickerIndex)
        Cells(4 + i, 2).Value = tickerVolumes(tickerIndex)
        Cells(4 + i, 3).Value = tickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex) - 1


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

The results for both the original and refactored code are the same, but the timer results are different.

##### Original Timer

Original Time for 2017

![Original Time for 2017](https://user-images.githubusercontent.com/84139177/124016073-da981600-d9aa-11eb-9548-4d8c099ef2fd.png)

Original Time for 2018

![Original Time for 2018](https://user-images.githubusercontent.com/84139177/124016090-dff56080-d9aa-11eb-97c5-63c05fc5b6ea.png)

##### Refactored Timer

Refactored Time for 2017

![Refactored Time for 2017](https://user-images.githubusercontent.com/84139177/124016245-0adfb480-d9ab-11eb-9fff-b2d1828186d4.png)

Refactored Time for 2018

![Refactored Time for 2018](https://user-images.githubusercontent.com/84139177/124016264-10d59580-d9ab-11eb-8391-01b08a5c32ad.png)

Based of the timer results, it looks like refactoring took about 6 to 7 seconds off the processing time.

## Summary
  ##### What are the advantages and disadvantages of refactoring code in general?
  Refactoring code can take the original code and make it run faster, decrease the memory use, and increase its overall functionality. Refactoring also encourages coders to revisit areas in code that could use better logic. Nonetheless, refactoring can create issues that were not present in the original code. These errors can be even more disastrous and time-consuming if the original code is not saved before refactoring attempts.
  
  ##### What are the advantages and disadvantages of the original and refactored VBA script?
  Refactoring in VBA is made easier by the fact that you can switch easily from one code to the next. This can make for an easy debugging process. However, refactoring VBA script can lead to issues involving incorrect syntax, overflow, or just a plain headache. It is worth taking time to consider if the code needs to be refactored before delving into the process.
