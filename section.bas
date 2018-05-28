Attribute VB_Name = "section"
Sub section()

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Set MyPlage = Range("G15:G1000")

    For Each cell In MyPlage
    
      If Not cell.Value = "" Then
      cell.Value = CDbl(cell.Value)
      End If
        Next
          

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
End Sub
