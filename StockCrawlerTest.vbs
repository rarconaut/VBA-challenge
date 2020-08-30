'Create a script that will loop through all the stocks for one year and output the following information.
Sub StockCrawlerTest()

Dim tickerStart As Variant
Dim tickerEnd As Variant
Dim tickerSymbol As String
Dim yearChange As Single
Dim totalVolume As Double
Dim stockOpen As Single
Dim stockClose As Variant
Dim percentChange As Single
Dim summaryRow As Integer

   
For Each ws In Worksheets
    summaryRow = 2
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    LastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
    
    ws.Cells(1, "I").Value = "Ticker"
    ws.Cells(1, "I").Columns.AutoFit
    
    ws.Cells(1, "J").Value = "Yearly Change ($)"
    ws.Cells(1, "J").Columns.AutoFit
    
    ws.Cells(1, "K").Value = "Percent Change"
    ws.Cells(1, "K").Columns.AutoFit
    
    ws.Cells(1, "L").Value = "Total Stock Volume"
    ws.Cells(1, "L").Columns.AutoFit
    
    ws.Cells(1, "P").Value = "Ticker"
    ws.Cells(1, "P").Columns.AutoFit
    
    ws.Cells(1, "Q").Value = "Value"
    ws.Cells(1, "Q").Columns.AutoFit
    
    ws.Cells(2, "O").Value = "Greatest % increase"
    ws.Cells(2, "O").Columns.AutoFit
    
    ws.Cells(3, "O").Value = "Greatest % decrease"
    ws.Cells(3, "O").Columns.AutoFit
    
    ws.Cells(4, "O").Value = "Greatest total volume"
    ws.Cells(4, "O").Columns.AutoFit
    
    For i = 2 To LastRow
    
    tickerSymbol = ws.Cells(i, "A").Value
            
  'The ticker symbol.
  '' Find this by looking for changes between the ticker name rows
        If tickerSymbol <> ws.Cells((i - 1), "A").Value Then
        tickerStart = i
        ''Makes the Ticker summary column = tickerSymbol
        ws.Cells(summaryRow, "I").Value = tickerSymbol

  'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
  '' Find this by subtracting the opening price from the first ticker row from the closing price from the last ticker row.
        stockOpen = ws.Cells(i, "C").Value
        
        ElseIf tickerSymbol <> ws.Cells((i + 1), "A").Value Then
        tickerEnd = i
        stockClose = ws.Cells(i, "F").Value
               
        yearChange = (stockClose - stockOpen)
        ws.Cells(summaryRow, "J").Value = yearChange
        
  'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
            If stockClose = 0 Then
            ws.Cells(summaryRow, "K").Value = "Collapse"
            Else
            percentChange = (yearChange / stockClose) * 100
            ws.Cells(summaryRow, "K").Value = percentChange
            End If
            
  'You should also have conditional formatting that will highlight positive change in green and negative change in red.
            If percentChange > 0 And percentChange <> "Collapse" Then
            ws.Cells(summaryRow, "K").Interior.Color = RGB(0, 255, 0)
            ElseIf percentChange < 0 Then
            ws.Cells(summaryRow, "K").Interior.Color = RGB(255, 0, 0)
            End If
            
  'The total stock volume of the stock.
  '' Sum of each stock's daily volumes
        For j = tickerStart To tickerEnd
        totalVolume = totalVolume + Cells(j, "G")
        ws.Cells(summaryRow, "L").Value = totalVolume
        Next j
        
  'Next row of summary table
        summaryRow = summaryRow + 1
        End If
    
    Next i


    
Dim greatestInc As Variant

Dim tickerInc As Variant

Dim greatestDec As Single

Dim tickerDec As String

Dim greatestVol As Double

Dim tickerVol As String

Dim greatest As Variant

ws.Cells(2, "Q").Value = greatestInc
ws.Cells(3, "Q").Value = greatestDec
ws.Cells(4, "Q").Value = greatestVol
ws.Cells(4, "Q").Columns.AutoFit

  ''grabbing the Greatest Increase/Decrease/Volume tickers
    For greatest = 2 To LastRow
    
    If (ws.Cells(greatest, "K").Value > greatestInc) Then
    greatestInc = ws.Cells(greatest, "K").Value
    ws.Cells(2, "Q").Value = greatestInc
    tickerInc = ws.Cells(greatest, "I").Value
    ws.Cells(2, "P").Value = tickerInc
    End If
    
    If (ws.Cells(greatest, "K").Value < greatestDec) Then
    greatestDec = ws.Cells(greatest, "K").Value
    ws.Cells(3, "Q").Value = greatestDec
    tickerInc = ws.Cells(greatest, "I").Value
    ws.Cells(3, "P").Value = tickerInc
    End If
    
    If (ws.Cells(greatest, "L").Value > greatestVol) Then
    greatestVol = ws.Cells(greatest, "L").Value
    ws.Cells(4, "Q").Value = greatestVol
    ws.Cells(4, "Q").Columns.AutoFit
    tickerInc = ws.Cells(greatest, "I").Value
    ws.Cells(4, "P").Value = tickerInc
    

    End If
    
    Next greatest
    
Next ws

End Sub
