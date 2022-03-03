# Stock-Analysis
Performing analysis on stock data to uncover trends and provide recommendations

Attached AllStockAnalysis Refactored macro to newly uploaded VBA and seen at the eend of this README

**Deliverable 1**
 
  _Prework_
   
The starter code "challenge_starter_code.vbs" was downloaded adn renamed to VBA_Challenge.vbs
A "Resources" folder was created to hold the run-time pop-up messages after running the refactored analysis 
The file "green_stocks.xlsm" was renamed to "VBA_Challenge.xlsm"
The "VBA_Challenge.vbs" script was added to the Microsoft Visual Basic editor 

    _Steps_
   
1a. A "tickerIndex" variable was created and set to zero before iterating over all the rows 
1b. Three out arrays were created with the following data types: "tickerVolumes" was created as a "Long" data type, "tickerStartingPrices" was created as a "Single" data type, and "tickerEndingPrices" was created as a "Single" data type
2a. A "for" loop was created to initialize the "tickerVolumes" to zero
2b. A "for" loop was created to loop over all the rows in the spreadsheet
3a. Inside the "for" loopa script was written that increases the current "tickerVolumes" variable and adds the ticker volume for the current stock ticker 
3b. An "if-then" statement was written to check if the current row is the first row with the selected "tickerIndex"; if so, it was assigned as the "tickerStartingPrices" variable
3c. An if-then statement was written to check if the current row is the last row with the selected "tickerIndex"; if so, it was assigned as the "tickerEndingPrices" variable 
3d. A script was written that increases the "tickerIndex" if the next row's ticker doesn't match the previous row's ticker
4. A "for" loop was used to loop through the aforementioned arrays to output the "Ticker", "TickerDailyVolume" and "Return" columns
The stock analysis was run to confirm the analysis from 2017 and 2018 matched the information from the module where the screenshots were saved into a separate "Resources" folder 

    _Refactored Macro_

1a. tickerIndex = 0

1b. Arrays:
  Dim tickerVolumes(12) As Long
  Dim tickerStartingPrices(12) As Single
  Dim tickerEndingPrices(12) As Single

2a. For i = 0 To 11
  tickerVolumes(i) = 0
  tickerStartingPrices(i) = 0
  tickerEndingPrices(i) = 0
  Next i

2b.
 For i = 2 To RowCount

3a.
  tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value

3b.
  If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
  tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
    End If
    
3c.
  If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
  tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
     End If

3d.
  If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
  tickerIndex = tickerIndex + 1
  End If

  Next i

4. For i = 0 To 11    
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = tickers(i)
    Cells(4 + i, 2).Value = tickerVolumes(i)
    Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    
Next i
End Sub

**Deliverable 2**

 _Section 1 Overview_

The purpose of this assignment was to use VBA to add an even more analytical power to Excel. In essence, the assignment involves utilizing VBA to create a script to better assist Steve, a somewhat savvy person in Excel with further marcos and analysis for adequate assessment and decision-making when reviewing stocks to invest for Steve's parents. 

 _Section 2 Results_

When reviewing the results, the stock performance between 2017 and 2018 was noticeably different; as well as the execution times. On one hand, comparing the year 2017 to 2018 out of the twelve stocks that were reviewed, 11 of them were "green" which means showed a positive return based on the macro. On the other hand, in 2018 10 ticker stocks were "red" which means depiected a negative return based on the macro. A deeper look into the stocks in 2017, shows that there was not just a slight positive return, but some ticker stocks such as DQ, ENPH, FSLR, and SEDG had a over 100% return which means the initial investment was at least doubled. As in a 100% positive return means the stock doubled in initial investment. 

The execution times of the original script and the refactored script are significantly different. In the original script, the run time for 2017 was 0.765625 and the run time for 2018 was 0.7617188. After the script had been refactored, the run time for 2017 was 0.1328125 and the rune time for 2018 was 0.1015625 (all screenshots of the difference in run times for each year are available in the resource folder). 

 _Section 3 Summary_

Clearly there are advantages and of refractoring code. Some advantages include making the code more "recognizeable" and easier to follow for other coders. This is because refractoring the code makes it "cleaner" and more oganized. This in turn also makes the code easier to read allowing it to be more simplified for support and updates, saving money and time in the future, and also and faster to maintain and "debug". Even with all these positives, like everything, there are some functional disadvantages. For example there would likely be a lot of re-testing for functionality. I encountered the same practice when refractoring, it took a good bit of time and I was required to go line by line to ensure correct functionality. Additionally, this was also a "smaller" script of code to refactor and there macros much larger that could take so much to complete it would not be a cost savings. 

These pros and cons apply to refractoring the original VBA script because as shown in the resources folder and mentioned above, after refractoring, the time it takes to run the code is shorter for both years of analysis (2017 and 2018). Moreover, for me personally, I was able to refractor the code in to a more organized and easy to read script for coders in the future in the event they were to review this specific macro. 



**AllStockAnalysis Refactored Code**

Sub AllStocksAnalysisRefactored_ChallengeAssignment()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("B1").Value = "All Stocks (" + yearValue + ")"
    Range("B1").Font.FontStyle = "Bold"
    
    'Create a header row
    Cells(3, 2).Value = "Ticker"
    Cells(3, 3).Value = "Total Daily Volume"
    Cells(3, 4).Value = "Return"
    Range("B3:D3").Interior.ColorIndex = 15

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
    
    
    'DELIVERABLE - REQUIREMENTS
    
    'The tickerIndex is set equal to zero before looping over the rows.
    
   For i = 0 To 11
       tickerIndex = tickers(i)
       
       
    'Arrays are created for tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices
    
    Dim tickerVolumes As Long
    Dim tickerStartingPrices As Single, tickerEndingPrices As Single
       
       
    'The tickerIndex is used to access the stock ticker index for the tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices arrays
    
       Worksheets(yearValue).Activate
       tickerVolumes = 0
       
       For j = 2 To RowCount
              
           
           'If the next row’s ticker doesn’t match, increase the tickerIndex.
           If Cells(j, 1).Value = tickerIndex Then
           
              
              'Increase volume for current ticker
              tickerVolumes = tickerVolumes + Cells(j, 8).Value
           
      
           End I
        
           If Cells(j - 1, 1).Value <> tickerIndex And Cells(j, 1).Value = tickerIndex Then

               'Store Starting Price Value
               tickerStartingPrices = Cells(j, 6).Value
               
           End If

           If Cells(j + 1, 1).Value <> tickerIndex And Cells(j, 1).Value = tickerIndex Then

               'Store Ending Price Value
               tickerEndingPrices = Cells(j, 6).Value
               

           End If
           
           
       Next j
     

           Worksheets("All Stocks Analysis").Activate
           
           Cells(4 + i, 2).Value = tickerIndex
           Cells(4 + i, 3).Value = tickerVolumes
           Cells(4 + i, 4).Value = tickerEndingPrices / tickerStartingPrices - 1
    
            'Fix % on return
            With Range("D4:D15")
                        .NumberFormat = "0.0%"
                        .Value = .Value
            End With
            

   Next i
 
   
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("B3:D3").Font.FontStyle = "Bold"
    Range("B3:D3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("C4:C15").NumberFormat = "#,##0"
    'Range("C4:C15").NumberFormat = "0.0%"
    Columns("C").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 4) > 0 Then
            
            Cells(i, 4).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 4).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub


