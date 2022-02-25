# Stock-Analysis
Performing analysis on stock data to uncover trends and provide recommendations
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
