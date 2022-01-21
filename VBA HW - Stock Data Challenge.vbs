Sub Stock_Data()

Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate

Dim total_vol As Double
Dim ticker_index As Integer
Dim count_ticker As Integer
Dim Opening_Price As Double
Dim Closing_Price As Double
Dim Yearly_Change As Double
Dim Percent_Change As Double

total_vol = 0
ticker_index = 2

count_ticker = 1

For i = 2 To 705714

    If (Cells(i + 1, 1).Value = Cells(i, 1).Value) Then
        total_vol = Cells(i, 7).Value + total_vol
        count_ticker = count_ticker + 1
    Else
        total_vol = Cells(i, 7).Value + total_vol
        Cells(ticker_index, 9).Value = Cells(i, 1).Value
        Cells(ticker_index, 12).Value = total_vol
        
        Opening_Price = Cells(i + 1 - count_ticker, 3)
        Closing_Price = Cells(count_ticker + 1, 6)
        Yearly_Change = Closing_Price - Opening_Price
        Cells(ticker_index, 10).Value = Yearly_Change
        
        If Cells(ticker_index, 10).Value > 0 Then
            Cells(ticker_index, 10).Interior.ColorIndex = 10
        Else
        Cells(ticker_index, 10).Interior.ColorIndex = 3
        End If
        
        Percent_Change = (Yearly_Change / Opening_Price)
        Cells(ticker_index, 11).Value = Percent_Change
        Cells(ticker_index, 11).NumberFormat = "0.00%"
        
        total_vol = 0
        ticker_index = ticker_index + 1
        count_ticker = 1
    End If
Next i