# stock-analysis
Analysis of Green Stocks data
# Outline of the project
### 
  The purpose of the All Stocks Analysis prject was to help Steve compare how various stocks fared over time and see if they had a positive or negative return. We are aiming to see if by refactoring the code, it will run more efficiently. 
# Results of the Analysis
###
  Many steps were taken to perform this analysis. The first step that was taken to refactor the code after the output sheet was activated, the header rows were created, and an array of all the tickers was initialized, was to create a ticker index and three output arrays. How this was accomplished can be seen below.  
  
      '1a) Create a ticker Index
       tickerIndex = 0
       
       '1b) Create three output arrays
        Dim tickerVolumes(11) As Long
        Dim tickerStartingPrices(11) As Single
        Dim tickerEndingPrices(11) As Single
        
  The next step was to create a loop that would loop over all of the rows in the worksheet, analyzing the values in each cell to see if it watched the selected ticker, and if it did, providing the starting and ending price values. The loop will then move to the next ticker until all stocks have been analyzied 
  
        ''2a) Create a for loop to initialize the tickerVolumes to zero.
        For i = 0 To 11
            tickerVolumes(i) = 0
            tickerStartingPrices(i) = 0
            tickerEndingPrices(i) = 0
        Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
        For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
                
        
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
             If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            End If
            
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            End If
            
            

            '3d Increase the tickerIndex.
                If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                    tickerIndex = tickerIndex + 1
                End If

    
    Next i
    
After all tickers have been analyzied, the next step was to output the data to the All Stocks Analysis worksheet. This worksheet was formatted to show which stocks had a positive return, shown in green, and which stocks had a negative return, shown in red. This was accomplished by using the following code to format the cells. 

      If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
When comparing the return results for 2017 and 2018, it can be seen that 2017 had only one stock, TERP, with a negative return and the rest of the stocks turned out positive. However, when looking at the returns for 2018, only two stocks, ENPH and RUN, return positive value and the rest were negative. The majority of stocks fared better and had higher return rates in 2017. 

After running the analysis on the original All Stocks Analysis code and the refactored code, I was able to see that the refactored code ran in a shorted amount of time. When I initially ran the All Stocks Analysis, the timer showed a time of over one second, an average of about 1.3 seconds. When I ran the same analysis on the refactored, for 2017 the code ran in about .37 second, seen in the image below. 
![VBA_Challenge_2017](https://github.com/aarce21/stock-analysis/blob/main/Resources/VBA_Challenge_2017.PNG)

When the analysis was initially run for 2018 with the original All Stocks Analysis code, the timer returned a run time of about 1.45 seconds. With the refactored code, 2018 returned a run time of about .35 seconds. 
![VBA_Challenge_2018](https://github.com/aarce21/stock-analysis/blob/main/Resources/VBA_Challenge_2018.PNG)

For both 2017 and 2018, the refactored code ran significantly quicker. 
# Summary 
## What are the advantages or disadvantages of refactoring code? 
  Although in this case refactoring the code allowed it to run more efficiently, there are a few disadvantages to refactoring code. One of which is that it can be a long and tedious process. Refactoring requires great attention to detail and typically a great deal of time. Refactoring code can come with spending a lot of time testing each individual part of the code to ensure there are no errors. 
  
  Refactoring code also has many advantages, one of which being that we will typically end up with a code that runs smoother and is easier to comprehend. It can also make it easier to debug a code if there is something incorrect. 
  
## How do these pros and cons apply to refactoring the original VBA script? 
  After refactoring the original VBA script, one of the advantages can clearly be seen by the decreased run time of the code. The refactor tidied up the code and allowed it to run more efficiently. The cons can also be seen by how long the process took to refactor the code, but it was worth it in the end. 
