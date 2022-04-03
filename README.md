# stock-analysis
Analysis of Green Stocks data
# Outline of the project
### 
  The purpose of the All Stocks Analysis prject was to help Steve compare how various stocks fared over time
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
    
After all tickers have been analyzied, the next step was to output the data to the All Stocks Analysis worksheet. 
# Summary 
## What are the advantages or disadvantages of refactoring code? 
## How do these pros and cons apply to refactoring the original VBA script? 
