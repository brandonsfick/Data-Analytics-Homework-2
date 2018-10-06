Sub Stock()
Dim stockName As String
Dim volumeTotal As Double
Dim lastRow As Variant
Dim currentSummaryRow As Integer
Dim I As Variant
Dim ws As Worksheet
Dim firstprice As Variant
Dim lastprice As Variant
Dim stockchange As Variant
Dim stockper As Variant
Dim j As Variant
Dim rng As Range
Dim dblmin As Variant
Dim dblminname As String
Dim dblmaxname As String
Dim dblmax As Variant
Dim volmax As Variant
Dim volmaxname As String


volumeTotal = 0
currentSummaryRow = 2
lastRow = Cells(Rows.Count, 1).End(xlUp).Row

For Each ws In ActiveWorkbook.Worksheets
ws.Activate
I = 2
volumeTotal = 0
currentSummaryRow = 2
lastRow = Cells(Rows.Count, 1).End(xlUp).Row
firstprice = Cells(I, 3).Value
Cells(1, 8).Value = "Ticker"
Cells(1, 9).Value = "Yearly Change"
Cells(1, 10).Value = "Percent Change"
Cells(1, 11).Value = "Total Stock Volume"
For I = 2 To lastRow
    If Cells(I + 1, 1).Value = Cells(I, 1).Value Then
        volumeTotal = volumeTotal + Cells(I, 7).Value
    Else
        stockName = Cells(I, 1).Value
        volumeTotal = volumeTotal + Cells(I, 7).Value
        lastprice = Cells(I, 6).Value
        stockchange = lastprice - firstprice
        stockper = (lastprice / firstprice) - 1
        stockper = Format(stockper, "Percent")
       
        Cells(currentSummaryRow, 8).Value = stockName
        Cells(currentSummaryRow, 9).Value = stockchange
        Cells(currentSummaryRow, 10).Value = stockper
        Cells(currentSummaryRow, 11).Value = volumeTotal
        If stockchange > 0 Then
            Cells(currentSummaryRow, 9).Interior.ColorIndex = 4
        Else
            Cells(currentSummaryRow, 9).Interior.ColorIndex = 3
        End If
        volumeTotal = 0
        currentSummaryRow = currentSummaryRow + 1
        If Cells(I + 1, 3) <> "" Then
            firstprice = Cells(I + 1, 3).Value
            If firstprice = 0 Then
                For j = I To lastRow
                    If firstprice = 0 Then
                    firstprice = Cells(j + 1, 3).Value
                    End If
                Next j
            End If
        End If
    End If

Next I
Cells(1, 15).Value = "Ticker"
Cells(1, 16).Value = "Value"
Cells(2, 14).Value = "Greatest % Increase"
Cells(3, 14).Value = "Greatest % Decrease"
Cells(4, 14).Value = "Greatest total volume"

dblmax = Cells(2, 10).Value
For a = 2 To lastRow

    If Cells(a, 10).Value > dblmax Then
    dblmax = Cells(a, 10).Value
    dblmaxname = Cells(a, 1).Value
    End If

Next a
Cells(2, 15).Value = dblmaxname
Cells(2, 16).Value = dblmax

dblmin = Cells(2, 10).Value

For b = 2 To lastRow
    
    If Cells(a, 10).Value < dblmin Then
    dblmin = Cells(a, 10).Value
    dblminname = Cells(a, 1).Value
    End If
   
Next b
Cells(3, 15).Value = dblminname
Cells(3, 16).Value = dblmin

volmax = Cells(2, 11).Value

For d = 2 To lastRow
    
    If Cells(a, 10).Value > volmax Then
    volmax = Cells(a, 10).Value
    volmaxname = Cells(a, 1).Value
    End If
   
Next d
Cells(4, 15).Value = volmaxname
Cells(4, 16).Value = volmax
Next ws

End Sub

