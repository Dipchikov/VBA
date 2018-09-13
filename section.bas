Attribute VB_Name = "section"
Sub section()
Dim lr As Long
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

lr = Range("A" & Rows.Count).End(xlUp).Row

On Error Resume Next
Set MyPlage = Range("G15:G" & lr)

    For Each cell In MyPlage
    
      If Not cell.Value = "" Then
      cell.Value = CDbl(cell.Value)
      End If
    Next
          

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
End Sub
