Attribute VB_Name = "renumber"
Sub renumber()
Dim lr As Long
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False


lr = Range("A" & Rows.Count).End(xlUp).Row

For i = 1 To 100
Set MyPlage = Range("B15:B" & lr)
For Each cell In MyPlage
        If cell.Value = i & ":" & i Then
        cell.Value = i
        cell.Font.ColorIndex = 3
        cell.Font.Bold = True
   End If
Next
  Set MyPlage = Range("E15:E" & lr)
    For Each cell In MyPlage
         If cell.Value = i & ":" & i Then
        cell.Value = i
        cell.Font.ColorIndex = 3
        cell.Font.Bold = True
   End If
  Next
Next i
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

