


Sub GreatestTest()
    
Dim greatestInc As Single

Dim tickerInc As String

Dim greatestDec As Single

Dim tickerDec As String

Dim greatestVol As Double

Dim tickerVol As String

Dim greatest As Variant

LastRow = Cells(Rows.Count, 1).End(xlUp).Row

  ''grabbing the Greatest Increase/Decrease/Volume tickers
For greatest = 2 To LastRow
    
    If (Cells(greatest, "K").Value > greatestInc) Then
    greatestInc = Cells(greatest, "K").Value
    Cells(2, "Q").Value = greatestInc
    tickerInc = Cells(greatest, "I").Value
    Cells(2, "P").Value = tickerInc
    End If
    
    If (Cells(greatest, "K").Value < greatestDec) Then
    greatestDec = Cells(greatest, "K").Value
    Cells(3, "Q").Value = greatestDec
    tickerInc = Cells(greatest, "I").Value
    Cells(3, "P").Value = tickerInc
    End If
    
    If (Cells(greatest, "L").Value > greatestVol) Then
    greatestVol = Cells(greatest, "L").Value
    Cells(4, "Q").Value = greatestVol
    tickerInc = Cells(greatest, "I").Value
    Cells(4, "P").Value = tickerInc
    End If
    
Next greatest
    
End Sub

