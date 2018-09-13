Attribute VB_Name = "TFM"
Sub tfm()

Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
'----------------- Shielded cable crossection   --------------------



Dim Shielded_cable As String
Set MyPlage = Range("L15:L1000")
Set myCell = MyPlage.Find(What:="Shielded cable", LookIn:=xlValues)
If Not myCell Is Nothing Then
Shielded_cable = InputBox("Please add cross-section for Shielded cable", , Range("G" & myCell.Row))
End If

Set MyPlage = Range("G15:G1000")
For Each cell In MyPlage
 If Not IsEmpty(cell.Value) And cell.Value <> Shielded_cable Then
If cell(1, 6).Value = "Shielded cable" Then
cell.Value = Shielded_cable
 cell.Font.ColorIndex = 3
 cell.Font.Bold = True
End If
End If
Next
Set MyPlage = Range("A15:A1000")
For Each cell In MyPlage
If Left(cell.Value, 3) = "TFM" And (cell(1, 2).Value = "13" Or cell(1, 2).Value = "14" Or cell(1, 2).Value = "39" Or cell(1, 2).Value = "40" Or cell(1, 2).Value = "41" Or cell(1, 2).Value = "42" Or cell(1, 2).Value = "43" Or cell(1, 2).Value = "44") And Left(cell(1, 4).Value, 3) = "XDC" And Not cell(1, 12).Value = "Shielded cable" Then
 tfm1 = InputBox("Please add cross-section of conductors between" & cell(1, 3) & " and " & cell(1, 6), "Wire jumper between " & cell(1, 3) & " and " & cell(1, 6), Shielded_cable)
        If tfm1 = vbOK Then
        cell(1, 7).Value = tfm1
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
End If
End If
If Left(cell.Value, 3) = "XDC" And (cell(1, 5).Value = "13" Or cell(1, 5).Value = "14" Or cell(1, 5).Value = "39" Or cell(1, 5).Value = "40" Or cell(1, 5).Value = "41" Or cell(1, 5).Value = "42" Or cell(1, 5).Value = "43" Or cell(1, 5).Value = "44") And Left(cell(1, 4).Value, 3) = "TFM" And Not cell(1, 12).Value = "Shielded cable" Then
tfm1 = InputBox("Please add cross-section of conductors between" & cell(1, 3) & " and " & cell(1, 6), "Wire jumper between " & cell(1, 3) & " and " & cell(1, 6), Shielded_cable)
        If tfm1 = vbOK Then
        cell(1, 7).Value = tfm1
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
End If
End If

Next

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
End Sub


