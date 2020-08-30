Sub GreatestTest()
    

Dim greatestInc As Variant
Dim greatestDec As Variant
Dim greatestVol As Double
Dim tickerInc As String
Dim tickerDec As String
Dim tickerVol As String

Dim greatestX As Variant


For Each ws In Worksheets

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

