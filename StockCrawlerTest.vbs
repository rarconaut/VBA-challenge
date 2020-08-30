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

   

    summaryRow = 2
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    LastColumn = Cells(1, Columns.Count).End(xlToLeft).Column
    
    Cells(1, "I").Value = "Ticker"
    Cells(1, "I").Columns.AutoFit
    
    Cells(1, "J").Value = "Yearly Change ($)"
    Cells(1, "J").Columns.AutoFit
    
    Cells(1, "K").Value = "Percent Change"
    Cells(1, "K").Columns.AutoFit
    
    Cells(1, "L").Value = "Total Stock Volume"
    Cells(1, "L").Columns.AutoFit
    
    Cells(1, "P").Value = "Ticker"
    Cells(1, "P").Columns.AutoFit
    
    Cells(1, "Q").Value = "Value"
    Cells(1, "Q").Columns.AutoFit
    
    Cells(2, "O").Value = "Greatest % increase"
    Cells(2, "O").Columns.AutoFit
    
    Cells(3, "O").Value = "Greatest % decrease"
    Cells(3, "O").Columns.AutoFit
    
    Cells(4, "O").Value = "Greatest total volume"
    Cells(4, "O").Columns.AutoFit
    
    For i = 2 To LastRow
    
    tickerSymbol = Cells(i, "A").Value
            
  'The ticker symbol.
  '' Find this by looking for changes between the ticker name rows
        If tickerSymbol <> Cells((i - 1), "A").Value Then
        tickerStart = i
        ''Makes the Ticker summary column = tickerSymbol
        Cells(summaryRow, "I").Value = tickerSymbol

  'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
  '' Find this by subtracting the opening price from the first ticker row from the closing price from the last ticker row.
        stockOpen = Cells(i, "C").Value
        
        ElseIf tickerSymbol <> Cells((i + 1), "A").Value Then
        tickerEnd = i
        stockClose = Cells(i, "F").Value
               
        yearChange = (stockClose - stockOpen)
        Cells(summaryRow, "J").Value = yearChange
        
  'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
            If stockClose = 0 Then
            Cells(summaryRow, "K").Value = "Collapse"
            Else
            percentChange = (yearChange / stockClose) * 100
            Cells(summaryRow, "K").Value = percentChange
            End If
            
  'You should also have conditional formatting that will highlight positive change in green and negative change in red.
            If percentChange > 0 And percentChange <> "Collapse" Then
            Cells(summaryRow, "K").Interior.Color = RGB(0, 255, 0)
            ElseIf percentChange < 0 Then
            Cells(summaryRow, "K").Interior.Color = RGB(255, 0, 0)
            End If
            
  'The total stock volume of the stock.
  '' Sum of each stock's daily volumes
        For j = tickerStart To tickerEnd
        totalVolume = totalVolume + Cells(j, "G")
        Cells(summaryRow, "L").Value = totalVolume
        Next j
        
  'Next row of summary table
        summaryRow = summaryRow + 1
        End If
    
    Next i


End Sub
