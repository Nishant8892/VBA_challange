Attribute VB_Name = "Module11"
Sub stockanalysis()
Dim ws As Worksheet
For Each ws In Worksheets
Dim totalstockvolume As Double
totalstockvolume = 0
Dim yearlychage As Double
yearlychange = 0
Dim percentchange As Double
percentchange = 0
Dim summary_table_row As Integer
summary_table_row = 2
Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
start_row = 2
For i = 2 To Lastrow
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
ticker = Cells(i, 1).Value
totalstockvolume = totalstockvolume + ws.Cells(i, 7).Value
yearlychange = ws.Cells(i, 6) - ws.Cells(start_row, 3)
If ws.Cells(start_row, 3) <> 0 Then
percentagechange = Round((yearlychange / ws.Cells(start_row, 3) * 100), 2)
End If
ws.Range("I" & summary_table_row).Value = ticker
ws.Range("J" & summary_table_row).Value = yearlychange
ws.Range("k" & summary_table_row).Value = percentagechange
ws.Range("L" & summary_table_row).Value = totalstockvolume
summary_table_row = summary_table_row + 1
totalstockvolume = 0
yearlychange = 0
percentagechange = 0
start_row = i + 1
Else
totalstockvolume = totalstockvolume + Cells(i, 7).Value
End If
Next i
ws.Cells(1, 9) = "ticker"
ws.Cells(1, 10) = "yearlychange"
ws.Cells(1, 11) = "percentagechange"
ws.Cells(1, 12) = "totalstockvolume"
Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To Lastrow
If ws.Cells(i, 10).Value > 0 Then
ws.Cells(i, 10).Interior.ColorIndex = 4
Else
End If
If ws.Cells(i, 10).Value < 0 Then
ws.Cells(i, 10).Interior.ColorIndex = 3
Else
End If
Next i
Next ws

End Sub
