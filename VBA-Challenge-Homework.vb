Sub TickerLoop()
Dim Ticker As String
Dim Change As Double
Dim Percent As Double
Dim Volume As Long
Dim Sumrow As Integer
Dim ws As Worksheet
Dim Open_Price As Double
Dim Close_Price As Double
    
For Each ws In Worksheets

        
Sumrow = 2
Volume = 0
    
    
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
    
    
lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
For i = 2 To lastrow

If Open_Price = 0 Then
Open_Price = ws.Cells(i, 3).Value


End If
        
    
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
Ticker = ws.Cells(i, 1).Value
Volume = Volume + ws.Cells(i, 7).Value
Close_Price = ws.Cells(i, 6).Value
Change = Close_Price - Open_Price
Percent = (Close_Price - Open_Price) / Open_Price
ws.Range("I" & Sumrow).Value = Ticker
ws.Range("L" & Sumrow).Value = Volume
ws.Range("J" & Sumrow).Value = Change
ws.Range("K" & Sumrow).Value = Percent
Sumrow = Sumrow + 1
Volume = 0


End If

ws.Range("K" & Sumrow).NumberFormat = "0.00%"
            
            
Next i

If ws.Cells(i, 10) >= 0 Then
ws.Cells(i, 10).Interior.ColorIndex = 4
Else
ws.Cells(i, 10).Interior.ColorIndex = 3
                
End If


Next ws