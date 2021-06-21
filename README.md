![VBA_Challenge_2017](https://user-images.githubusercontent.com/85076259/122702485-b4d17b00-d20c-11eb-9b41-39eecd8c5148.PNG)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/85076259/122702488-b7cc6b80-d20c-11eb-8d8e-918449d9e73a.PNG)
# stock-analysis

Stock Analysis Overview 

Purpose
In this project we edited and refactors Stock Market Data with VBA to help Steve analyze data more efficiently. Steve wanted to know if the stocks were worth investing in in the long run based on the data from 2017-2018. 

Analysis
Before refactoring the code, I copied the code that was provided in the input box. Then I filled in the specified portions.
'1a) Create a ticker Index
   tickerindex = 0
    
    

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerindex = tickers(i)
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    Next i
    
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerindex) = tickerVolumes(tickerindex) + Cells(i, 6).Value
        
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
         If Cells(i, 1).Value = tickers(tickerindex) And Cells(i - 1, 1).Value <> tickers(tickerindex) Then
            tickerStartingPrices(tickerindex) = Cells(i, 6).Value
        End If
        
         
            
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
          If Cells(i, 1).Value = tickers(tickerindex) And Cells(i + 1, 1).Value <> tickers(tickerindex) Then
        tickerEndingPrices(tickerindex) = Cells(i, 6).Value
     End If
            

            '3d Increase the tickerIndex.
            If Cells(i, 1).Value = tickers(tickerindex) And Cells(i + 1, 1).Value <> tickers(tickerindex) Then
            tickerindex = tickerindex + 1
            
            End If
            
       Next i
       
        
            
        'End If
    
   
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
    
    Refactoring Code: Pros and Cons
    
    Refactoring the code helps us stay more organized and allows the code to be cleaner. A few advatanges of working with cleaner code is design, debugging and faster programming. It is also easier to read. However, I noticed the disadvantage of refactoring our code is having a data set that is too large or not having the proper test cases for the existing codes and pose some risks to our code. I noticed that the biggest advantage of refactoring our code in this eample was how fast the run times were when we ran it. Attached are the run times for both 2018 and 2017. 
