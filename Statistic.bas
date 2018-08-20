Attribute VB_Name = "Statistic"
Sub Statistic()

Set Data = Sheets("Wiring table")
Set Final = Sheets("Statistic")
Dim irow As Long
Dim rng As Range

irow = Final.Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row

Final.Cells(irow, 1).Value = Date
Final.Cells(irow, 2).Value = Data.Range("G1").Value
Final.Cells(irow, 3).Value = Data.Range("B1").Value
Final.Cells(irow, 4).Value = Data.Range("H10").Value

Set rng = Final.Range("A2:D" & irow)
With rng.Borders
        .LineStyle = xlContinuous
        '.Color = vbRed
        .Weight = xlThin
End With
    

irow = irow + 1


End Sub
