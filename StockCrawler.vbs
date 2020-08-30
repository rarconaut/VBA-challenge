

'Create a script that will loop through all the stocks for one year
Sub StockCrawler()

Dim tickerStart As Variant
Dim tickerEnd As Variant
Dim tickerSymbol As String
Dim yearChange As Single
Dim totalVolume As Double
Dim stockOpen As Single
Dim stockClose As Single
Dim percentChange As Variant
Dim summaryRow As Integer

'' Start the loop for going through each worksheet
For Each ws In Worksheets

    '' Set row count for new output table
    summaryRow = 2
    
    '' Count rows for each worksheet/year
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    '' Add header for new output table
    ws.Cells(1, "I").Value = "Ticker"
    ws.Cells(1, "I").Columns.AutoFit
    
    ws.Cells(1, "J").Value = "Yearly Change ($)"
    ws.Cells(1, "J").Columns.AutoFit
    
    ws.Cells(1, "K").Value = "Percent Change"
    ws.Cells(1, "K").Columns.AutoFit
    
    ws.Cells(1, "L").Value = "Total Stock Volume"
    ws.Cells(1, "L").Columns.AutoFit
    
    
    
  '' Crawl through the stock data row by row, from row2 to the last row counted
    For i = 2 To LastRow
       
  'Output the following information:
  '--The ticker symbol.
    tickerSymbol = ws.Cells(i, "A").Value
    
        '' Find this by looking for changes between the ticker row names
        If tickerSymbol <> ws.Cells((i - 1), "A").Value Then
        tickerStart = i
        
        ''Makes the Ticker summary column = to the tickerSymbol value
        ws.Cells(summaryRow, "I").Value = tickerSymbol

  '--Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
  '' Find this by subtracting the opening price from the first ticker row from the closing price from the last ticker row.
        stockOpen = ws.Cells(i, "C").Value
        
        ElseIf tickerSymbol <> ws.Cells((i + 1), "A").Value Then
        tickerEnd = i
        stockClose = ws.Cells(i, "F").Value
               
        yearChange = (stockClose - stockOpen)
        ws.Cells(summaryRow, "J").Value = yearChange
        
  '--The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
  '' Need an extra If statement to avoid the 'divide by zero' case
            If stockClose = 0 Then
            ws.Cells(summaryRow, "K").Value = "Collapse"
            Else
            percentChange = (yearChange / stockClose) * 100
            ws.Cells(summaryRow, "K").Value = percentChange
            End If
            
  'You should also have conditional formatting that will highlight positive change in green and negative change in red.
            If percentChange <> "Collapse" And percentChange > 0 Then
            ws.Cells(summaryRow, "K").Interior.Color = RGB(0, 255, 0)
            ElseIf percentChange <> "Collapse" And percentChange < 0 Then
            ws.Cells(summaryRow, "K").Interior.Color = RGB(255, 0, 0)
            ElseIf percentChange = "Collapse" Then
            ws.Cells(summaryRow, "K").Interior.Color = RGB(0, 0, 0)
            End If
            
  '--The total stock volume of the stock.
  '' Sum of each stock's daily volumes
        For j = tickerStart To tickerEnd
        totalVolume = totalVolume + Cells(j, "G")
        ws.Cells(summaryRow, "L").Value = totalVolume
        Next j
        
  '' Next row of summary table
        summaryRow = summaryRow + 1
        End If
    
    Next i
    




'###Challenge Solutions
'Make the appropriate adjustments to your VBA script that will allow it to run on every worksheet,
'i.e., every year, just by running the VBA script once.

    

Dim greatestInc As Variant
Dim greatestDec As Variant
Dim greatestVol As Double
Dim tickerInc As String
Dim tickerDec As String
Dim tickerVol As String

Dim greatestX As Variant


''creating headers for summary table
ws.Cells(2, "O").Value = "Greatest % increase"
ws.Cells(3, "O").Value = "Greatest % decrease"
ws.Cells(4, "O").Value = "Greatest total volume"
ws.Cells(4, "O").Columns.AutoFit
    
ws.Cells(1, "P").Value = "Ticker"
ws.Cells(1, "P").Columns.AutoFit
    
ws.Cells(1, "Q").Value = "Value"


LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

  ''grabbing the Greatest Increase/Decrease/Volume and tickers
    For greatestX = 2 To LastRow
    
    greatestInc = ws.Cells(2, "Q").Value
    greatestDec = ws.Cells(3, "Q").Value
    greatestVol = ws.Cells(4, "Q").Value
    
  ''grabbing the Greatest Increase
    If ws.Cells(greatestX, "K").Value <> "Collapse" And ws.Cells(greatestX, "K").Value > greatestInc Then
    greatestInc = ws.Cells(greatestX, "K").Value
    ws.Cells(2, "Q").Value = greatestInc
  ''grabbing the Greatest Increase ticker
    tickerInc = ws.Cells(greatestX, "I").Value
    ws.Cells(2, "P").Value = tickerInc
    End If
    
  ''grabbing the Greatest Decrease
    If ws.Cells(greatestX, "K").Value <> "Collapse" And ws.Cells(greatestX, "K").Value < greatestDec Then
    greatestDec = ws.Cells(greatestX, "K").Value
    ws.Cells(3, "Q").Value = greatestDec
  ''grabbing the Greatest Decrease ticker
    tickerInc = ws.Cells(greatestX, "I").Value
    ws.Cells(3, "P").Value = tickerInc
    End If
    
  ''grabbing the Greatest Volume
    If ws.Cells(greatestX, "L").Value > greatestVol Then
    greatestVol = ws.Cells(greatestX, "L").Value
    ws.Cells(4, "Q").Value = greatestVol
    ws.Cells(4, "Q").Columns.AutoFit
    
  ''grabbing the Greatest Volume ticker
    tickerInc = ws.Cells(greatestX, "I").Value
    ws.Cells(4, "P").Value = tickerInc
    End If
    
    Next greatestX
    
Next ws

End Sub


