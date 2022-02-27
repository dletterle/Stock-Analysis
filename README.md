# Stock-Analysis
Performing analysis on stock data to uncover trends and provide recommendations

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

    _Macro_
    

**Deliverable 2**

 _Section 1 Overview_

The purpose of this assignment was to use VBA to add an even more analytical power to Excel. In essence, the assignment involves utilizing VBA to create a script to better assist Steve, a somewhat savvy person in Excel with further marcos and analysis for adequate assessment and decision-making when reviewing stocks to invest for Steve's parents. 

 _Section 2 Results_

When reviewing the results, the stock performance between 2017 and 2018 was noticeably different; as well as the execution times. On one hand, comparing the year 2017 to 2018 out of the twelve stocks that were reviewed, 11 of them were "green" which means showed a positive return based on the macro. On the other hand, in 2018 10 ticker stocks were "red" which means depiected a negative return based on the macro. A deeper look into the stocks in 2017, shows that there was not just a slight positive return, but some ticker stocks such as DQ, ENPH, FSLR, and SEDG had a over 100% return which means the initial investment was at least doubled. As in a 100% positive return means the stock doubled in initial investment. 

The execution times of the original script and the refactored script are significantly different. In the original script, the run time for 2017 was 0.765625 and the run time for 2018 was 0.7617188. After the script had been refactored, the run time for 2017 was 0.1328125 and the rune time for 2018 was 0.1015625 (all screenshots of the difference in run times for each year are available in the resource folder). 

 _Section 3 Summary_

Clearly there are advantages and disadvanteges of refractoring code. Some advantages include making the code more "recognizeable" and easier to follow for other coders. This is because refractoring the code makes it "cleaner" and more oganized. This in turn also makes the code easier to read allowing it to be more simplified for support and updates, saving money and time in the future, and also and faster to maintain and "debug". Even with all these positives, like everything, there are some functional disadvantages. For example there would likely be a lot of re-testing for functionality. I encountered the same practice when refractoring, it took a good bit of time and I was required to go line by line to ensure correct functionality. Additionally, this was also a "smaller" script of code to refactor and there macros much larger that could take so much to complete it would not be a cost savings. 



Dataset was downloaded
Dataset was saved as macro "xslm"
General subroutine MacroCheck was ran to make sure VBA is running correctly 
Test Message ran with "Hello World!" in message box to check for correct syntax and macro enabled 
Original file uploaded to GitHub
Created data headers utilizing Range() and Cell() methods
Practiced cell and range method by switching cell(3,1) to range(A3)
Uploaded file to GitHub
Added a row of white space 
Created loop to calculate total volume and year 
Upload file to GitHub
Error in workbook created new excel file
Uploaded new excel file to GitHub
Created and sited new pattern with expert code
Uploaded file to GitHub
Re-wrote code to include logical statements in order to calcuate the yearly return; found the rows where DQ stock started and ended then tabulated the rows of data in between
Created Practice Workbook with the following code
Sub Analysis()
   Worksheets("Practice").Activate
'Make a list of square numbers
For i = 1 to 10
    Cells(1, i).Value = i * i
Next i
End Sub
Uploaded new code to GitHub
Created new worksheet called "All Stocks Analysis" to run analysis on all stocks in 2018, created a module to match the parameters of the new worksheet in VBA
Created a nested poll loop to put 1 in cells A1 to J1; A10 to J10 with the following code
Sub nestedLoopFor()
  Worksheets("PracticeLoop").Activate
  Dim r As Integer, c As Integer
For r = 1 To 10
For c = 1 To 10
    Cells(r, c).Value = 1
    Next
 Next
 End Sub
For skill drill 2.4.2 created a lengthy macro to run column by column alternating changing the color between the cells based on the values inside. Values were either positive or negative "1" to give the macro a value to run. The -1 cell was coded to be red and the 1 cell was coded to be black to give a checkerboard appearance
