'# VBA-challenge
'Homework #2 for Data Analytics Bootcamp

Sub Stocks()
Dim i, j, vol As Variant
Dim ticker As String
Dim max2, min2, diff, percentchange As Double
Dim lastrow, ws_count, y As Integer
Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets
ws.Activate

min2 = CDbl(Range("C2").Value)

Range("I1") = "Ticker"
Range("J1") = "Yearly Change"
Range("K1") = "Percent Change"
Range("L1") = "Total Stock Volume"


For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
        max2 = Cells(i, 6).Value
        diff = max2 - min2
        
        If diff = 0 Or min2 = 0 Then
            percentchange = 0
        Else
            percentchange = diff / min2
        End If
        
        min2 = Cells(i + 1, 3).Value
        ticker = Cells(i, 1)
        Cells(Cells(Rows.Count, 9).End(xlUp).Row + 1, 9) = ticker
        Cells(Cells(Rows.Count, 10).End(xlUp).Row + 1, 10).Value = diff
        Cells(Cells(Rows.Count, 11).End(xlUp).Row + 1, 11).NumberFormat = "0.00%"
        Cells(Cells(Rows.Count, 11).End(xlUp).Row + 1, 11) = percentchange
        vol = vol + Cells(i, 7).Value
        Cells(Cells(Rows.Count, 12).End(xlUp).Row + 1, 12).Value = vol
        vol = 0
    Else
        vol = vol + Cells(i, 7).Value
    End If
    
Next i

For j = 2 To Cells(Rows.Count, 9).End(xlUp).Row
    If Cells(j, 10) < 0 Then
        Cells(j, 10).Interior.ColorIndex = 3
    Else
        Cells(j, 10).Interior.ColorIndex = 4
    End If
Next j

Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
Range("O2").Value = "Greatest % increase"
Range("O3").Value = "Greatest % decrease"
Range("O4").Value = "Greatest Total Value"
lastrow = Cells(Rows.Count, "A").End(xlUp).Row

Range("Q2") = WorksheetFunction.Max(Range("K2:K" & lastrow))
Range("Q3") = WorksheetFunction.Min(Range("K2:K" & lastrow))
Range("Q4") = WorksheetFunction.Max(Range("L2:L" & lastrow))
 
increase_num = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & lastrow)), Range("K2:K" & lastrow), 0)
decrease_num = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & lastrow)), Range("K2:K" & lastrow), 0)
volume_num = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & lastrow)), Range("L2:L" & lastrow), 0)
    Range("P2") = Cells(increase_num + 1, 9).Value
    Cells(2, 17).NumberFormat = "0.00%"
    Range("P3") = Cells(decrease_num + 1, 9).Value
    Cells(3, 17).NumberFormat = "0.00%"
    Range("P4") = Cells(volume_num + 1, 9).Value

Next ws
End Sub
