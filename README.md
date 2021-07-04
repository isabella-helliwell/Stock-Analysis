# VBA_Challenge

## 1. Project Scope
      Part1:
      The scope of this project is to compare the stockmarket for predefined stocks for years 2017 and 2018
      writing a VBA macro in Excel.
      Part 2:
      The second part of this project is to optimize the VBA code and Runtime
      
### 1.1 Stockmarket Analysis and Results
      The analysis are based on obtaining the yearly return for the predefined stocks.
      The formula used is (endingPrice / startingPrice - 1) for each year and stock.
      Results are shown in Table 1 and Table 2.
 
 Table 1. Yearly return for year 2017
 
![image](https://user-images.githubusercontent.com/85843030/124384753-c69d2e80-dca0-11eb-9eba-b0d7f01f6c8b.png)




Table2. Yearly return for year 2018

![image](https://user-images.githubusercontent.com/85843030/124384801-fea47180-dca0-11eb-861f-8f5bf24c4c5e.png)

Overal, the stocks did much better in 2017 than 2018.

### 1.2 VBA Code
      The overal layout of the VBA code consists of a suberoutine name, and a collection of steps and instructions.
      The name of the Subroutine is set to "VBA_Challenge", and the instructions given is to calculate the yearly return for stocks.
      In the first part of the macro, variable declarations are made to assign the correct data types.
      In the main part of the macro, the actual calculations are carried out utilizing nested foor loops.
      In the last part of the macro, the results are populated into the appropiate cells and sheets.
      There are also some formatting and macro runtime recorded.
      
#### 1.2.1 VBA Code Explained

      The following code is used for the first part of the macro:
##### Sub VBA_Challenge()
##### Dim startTime As Single
##### Dim endtime As Single
##### Dim tickers(12) As String
##### Dim startingPrice As Single
##### Dim endingprice As Single
##### Dim totalvolume As Long
      The above code, is declaring some variables using the following format: [DIM variablename AS DATATYPE],
      
##### yearvalue = InputBox("What year would you like to run the analysis on")
      -MessageBox will ask the user to input year for the Stocks to be analysed
##### RowCount = Sheets(yearvalue).Cells(Rows.Count, "A").End(xlUp).Row
      - RowCount will open up the Excel tab with the same year as the user year input and select column A,
        and scroll up to the first row with a non blank cell in column A. This will provided how many Rows there is in the sheet opened.
##### starttimer = Timer
      -Starts timing the macro
##### Worksheets(yearvalue).Activate
      -activates the tab in Excel that the user input year chose in the previous steps
      
      
      
      
      
      
