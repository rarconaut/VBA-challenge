

Sub WSLooper()

Dim tickerStart As Variant
Dim tickerEnd As Variant
Dim tickerSymbol As String
Dim yearChange As Single
Dim totalVolume As Double
Dim stockOpen As Single
Dim stockClose As Single
Dim percentChange As Variant
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
            If percentChange <> "Collapse" And percentChange > 0 Then
            ws.Cells(summaryRow, "K").Interior.Color = RGB(0, 255, 0)
            ElseIf percentChange <> "Collapse" And percentChange < 0 Then
            ws.Cells(summaryRow, "K").Interior.Color = RGB(255, 0, 0)
            ElseIf percentChange = "Collapse" Then
            ws.Cells(summaryRow, "K").Interior.Color = RGB(0, 0, 0)
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
    
Next ws

End Sub

