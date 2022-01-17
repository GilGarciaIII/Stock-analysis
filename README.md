# Stock-analysis
OverView of Project Explaine the purpose of this analysis. 
    The overall purpose of this analysis, was to analyis Steve's 12 different stocks data using excel and VBA. The reason behind analyis of Steve's stock is to determine if they are worth investing in. 
    
Results

    The Results in the analysis deals with 12 different stocks, and their history. I Examand the Value of each stock and the date of the stock was issued and the difference between the opening price and the closing price, also the highest and lowest stock prices. End goal was to see what the Value, daily volume, and the return on each of the 12 stocks are, if there was even any return at all.  
    
    
Summary 

    What are the advantages or disadvantages of refactoring code? The advantage of refactoring the code from the orginal code is to make it more orginized. Clean code is easier to read and understand, also more efficiant and faster programing using shortcuts. Also easier to understand what is going on and what the end goal is. The disadvantage is sometimes you cannot refactor code at all, due to an application being to big or not having the correct refactoring code to reach the goal of your analysis.
    How do these pros and cons apply to refactoring the original VBA script?
       The speed it took for the refactor code to analyis the stock. The macro run rate was slower than the orginal VBA script.
        
Code used:

    '1a) Create a ticker Index
tickerIndex = 0

'1b) Create three output arrays
Dim tickerVolumes(12) As Long
Dim tickerStartingPrices(12) As Single
Dim tickerEndingPrices(12) As Single

''2a) Create a for loop to initialize the tickerVolumes to zero.
' If the next row’s ticker doesn’t match, increase the tickerIndex.
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
    'If  Then
     If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
     End If

        '3d Increase the tickerIndex.
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

   
