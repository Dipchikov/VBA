Attribute VB_Name = "Statistic"
Sub Statistic()
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False

Set Data = Sheets("Wiring table")
Set Final = Sheets("Statistic")
Dim irow As Long
Dim rng As Range, c As Range
Dim lr As Long
'------errors--counter------------------
Dim Crosssection As Long, lastrow As Long
Dim Colour As Long
Dim connection As Long
Dim designation As Long
Dim Designation2 As Long
'Dim rng As Range, c As Range
'Find the last row
lastrow = Data.Range("A" & Rows.Count).End(xlUp).Row
'Set the range you want to search through
Set rng = Data.Range("G15:G" & lastrow)

'Iterate through each cell in the range
For Each c In rng
    'If the interior color is 6 (standard yellow), not blank and not a number
    If c.Font.ColorIndex = 3 Then
        'Add 1 to numbers
        Crosssection = Crosssection + 1
    End If
Next c

Set rng = Data.Range("H15:H" & lastrow)

For Each c In rng
    'If the interior color is 6 (standard yellow), not blank and not a number
    If c.Font.ColorIndex = 3 Then
        'Add 1 to numbers
        Colour = Colour + 1
    End If
Next c

Set rng = Data.Range("I15:i" & lastrow)

For Each c In rng
    'If the interior color is 6 (standard yellow), not blank and not a number
    If c.Font.ColorIndex = 3 Then
        'Add 1 to numbers
        connection = connection + 1
    End If
Next c

Set rng = Data.Range("A15:F" & lastrow)

For Each c In rng
    'If the interior color is 6 (standard yellow), not blank and not a number
    If c.Font.ColorIndex = 3 Then
        'Add 1 to numbers
        designation = designation + 1
    End If
Next c
Set rng = Data.Range("I15:i" & lastrow)

For Each c In rng
    'If the interior color is 6 (standard yellow), not blank and not a number
    If c.Font.ColorIndex = 6 Then
        'Add 1 to numbers
        Designation2 = Designation2 + 1
    End If
Next c





'If Designation2 + designation + Crosssection + Colour + connection > 0 Then ' remove this conndition
Final.Select
irow = Final.Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row
lr = Final.Range("C" & Rows.Count).End(xlUp).Row
Set MyPlage = Range("C2:C" & lr)
For Each cell In MyPlage
If cell.Value = Data.Range("B1").Value Then

answer = MsgBox("This schematic already exist!!!" & vbNewLine & "If you want to replace then press - Yes", vbYesNo + vbQuestion + vbDefaultButton2, "Comax table")
    If answer = vbYes Then
    Range("A" & cell.Row).Value = Date
    Range("B" & cell.Row).Value = Data.Range("G1").Value
    Range("C" & cell.Row).Value = Data.Range("B1").Value
    Range("D" & cell.Row).Value = Designation2 + designation
    Range("E" & cell.Row).Value = Crosssection
    Range("F" & cell.Row).Value = connection
    Range("H" & cell.Row).Value = Designation2 + designation + Crosssection + Colour + connection
    Range("I" & cell.Row).Value = Data.Range("F10").Value
    Range("G" & cell.Row).Value = MonthName(Month(Date))
    Exit Sub
Else
Exit Sub
End If
    
End If
Next
'If Cell.Value <> Data.Range("B1").Value Then
Final.Cells(irow, 1).Value = Date
Final.Cells(irow, 2).Value = Data.Range("G1").Value
Final.Cells(irow, 3).Value = Data.Range("B1").Value
Final.Cells(irow, 4).Value = Designation2 + designation
Final.Cells(irow, 5).Value = Crosssection
Final.Cells(irow, 6).Value = Colour
Final.Cells(irow, 7).Value = connection
Final.Cells(irow, 8).Value = Designation2 + designation + Crosssection + Colour + connection
Final.Cells(irow, 9).Value = Data.Range("F10").Value
Final.Cells(irow, 10).Value = MonthName(Month(Date))

Set rng = Final.Range("A2:J" & irow)
With rng.Borders
        .LineStyle = xlContinuous
        '.Color = vbRed
        .Weight = xlThin
End With
'End If


irow = irow + 1




Data.Select
'End If
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True

End Sub
