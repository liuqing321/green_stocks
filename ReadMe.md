
## Overview of Project

To help his parent to diversify their investment, Steve is investigating some stocks from green energy industry. 
Steve wants to expand the dataset to include the entire stock market over the last few years. Although the current code works well for a dozen stocks, it may take a long time to execute if larger dataset was involved. 
 
## Purpose

The purpose of this analysis is to refactor the given code, and improve the efficiency of the give code. So Steve could apply this code to the data from the entire stock market over the past few years. 


## Results

### Analysis for Stock Performance 

Comparing the result in 2017 to the result in 2018, we could see that the performance for the entire stock market was much better than it was in 2018. Only one stock had a negative return rate in 2017. 

Most of the investments in 2017 were profitable. In 2018, only two stocks still had positive return rates,RUN and ENPH. 

Steve's parents could consider buying some shares from RUN and ENPH. 

### Comparison between the original code and the refactored code

From the images above, we can see that the refactored code is more efficient than the original one. The refactored code running much more faster than the original one. 

### Refactored code 

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

    For j = 2 To RowCount
   
   '3a) Increase volume for current ticker
    
    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
    
   
   '3b) Check if the current row is the first row with the selected tickerIndex.
  
    
    If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j - 1, 1).Value <> tickers(tickerIndex) Then
        
        tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
    
    End If
    
   '3c) check if the current row is the last row with the selected ticker
    
    
     If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
        tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
     End If

   '3d Increase the tickerIndex.

         If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
        End If


        
   '3c) check if the current row is the last row with the selected ticker

         'If the next rows ticker doesn't match, increase the tickerIndex.
        
            
            If Cells(j + 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex) Then
             
            tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
            
            End If
            

   '3d Increase the tickerIndex.
      
            
        If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
        tickerIndex = tickerIndex + 1
        End If
        Next j
    
   '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    
    Worksheets("All Stocks Analysis").Activate
    For i = 0 To 11
    
    Cells(4 + i, 1).Value = tickers(i)
    Cells(4 + i, 2).Value = tickerVolumes(i)

## Summary 

### The advantages and disadvantages of refactoring code 

- Advantages
  
    Refactoring code could make the code more efficient by taking fewer steps, using less memeory, or improving the logic of the code to make it easier for future users to read. 

- Disadvantages 

    Refactoring requires extracting software system structure, data models, and intra-application dependencies to get back knowledge of an existing software system. Sometimes the person who are refactoring the code doesn't have the accurate knowledge for the current state of a system or about the design decisions made by the previous developer. And refactoring activities might deteriorate the structure architechture of a system, and such deteriorate could lead to re-development of the system. 

### The advantages and disadvantages of the original and refactored VBA script

  Before refactoring, there is only one output array in the script. So it took longer to loop over all the rows.

  After refactoring, the script ran faster than the original one, which means the refacored script is more efficient. Because there are 3 arrays in the current script now. We could applied the refactored script to a dataset that is much more larger than the current dataset. 

  additionally, the refactored script became more readable and organized. It will be easier to locate and fix the bug. 

  However, the refactored script is more complex. If someone wants to modify the script and apply it to another dataset, or use the the script under different scenarios, he or she will spend more time on changing the elements in the script. The refactored script might not as adaptable as the original one.  



    

 

