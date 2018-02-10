Sub multi_year_stock()

For Each ws In Worksheets

Dim Ticker As String
Dim Total_Stock_Volume As Double
Dim stock_vol_sum As Integer

stock_vol_sum = 2

For i = 2 To 797711


If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

'set Ticker name
    Ticker = ws.Cells(i, 1).Value

'add to the Total Stock Volume per Ticker
    Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value

'print Ticker header into I1
    ws.Range("I1").Value = "Ticker"

'print Total Stock Volume header into J1
    ws.Range("J1").Value = "Total Stock Volume"

'print Ticker name into col I
    ws.Range("I" & stock_vol_sum).Value = Ticker

'print Total_Stock_Volume into col J
    ws.Range("J" & stock_vol_sum).Value = Total_Stock_Volume

    stock_vol_sum = stock_vol_sum + 1

'reset total stock volume counter
    Total_Stock_Volume = 0

Else

    Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value

End If

    
Next i
Next ws
End Sub
