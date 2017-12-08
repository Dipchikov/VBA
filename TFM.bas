Attribute VB_Name = "TFM"
Sub tfm()

'----------------- Shielded cable crossection   --------------------

Dim Shielded_cable As String

Set myplage = Range("L15:L1000")
Set myCell = myplage.Find(What:="Shielded cable", LookIn:=xlValues)
If Not myCell Is Nothing Then
Shielded_cable = InputBox("Please add cross-section for Shielded cable", , "0,8")
End If

Set myplage = Range("A15:A1000")
For Each cell In myplage
If Left(cell.Value, 3) = "TFM" And (cell(1, 2).Value = "13" Or cell(1, 2).Value = "14" Or cell(1, 2).Value = "39" Or cell(1, 2).Value = "40" Or cell(1, 2).Value = "41" Or cell(1, 2).Value = "42" Or cell(1, 2).Value = "43" Or cell(1, 2).Value = "44") And Left(cell(1, 4).Value, 3) = "XDC" And Not cell(1, 12).Value = "Shielded cable" Then
 tfm1 = InputBox("Please add cross-section of conductors between" & cell(1, 3) & " and " & cell(1, 6), "Wire jumper between " & cell(1, 3) & " and " & cell(1, 6), Shielded_cable)
        cell(1, 7).Value = tfm1
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
End If

If Left(cell.Value, 3) = "XDC" And (cell(1, 5).Value = "13" Or cell(1, 5).Value = "14" Or cell(1, 5).Value = "39" Or cell(1, 5).Value = "40" Or cell(1, 5).Value = "41" Or cell(1, 5).Value = "42" Or cell(1, 5).Value = "43" Or cell(1, 5).Value = "44") And Left(cell(1, 4).Value, 3) = "TFM" And Not cell(1, 12).Value = "Shielded cable" Then
tfm1 = InputBox("Please add cross-section of conductors between" & cell(1, 3) & " and " & cell(1, 6), "Wire jumper between " & cell(1, 3) & " and " & cell(1, 6), Shielded_cable)
        cell(1, 7).Value = tfm1
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True

End If

Next
End Sub
