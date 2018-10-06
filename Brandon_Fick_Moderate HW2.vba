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
Next ws

End Sub

