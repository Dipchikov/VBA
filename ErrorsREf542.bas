Attribute VB_Name = "ErrorsREf542"
Sub ErrorsREf542()
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
Set MyPlage = Range("A15:A1000")
  For Each cell In MyPlage


                    If Left(cell.Value, 2) = "AA" And cell(1, 13).Value > 1 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
                    If Left(cell.Value, 2) = "AA" And cell(1, 13).Value <= 1 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If


Next

Set MyPlage = Range("D15:D1000")
  For Each cell In MyPlage


                    If Left(cell.Value, 2) = "AA" And cell(1, 11).Value > 1 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
                    If Left(cell.Value, 2) = "AA" And cell(1, 11).Value <= 1 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If


Next
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
