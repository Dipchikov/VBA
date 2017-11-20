Attribute VB_Name = "ErrorsRefs"
Sub ErrorsRefs()
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
Set MyPlage = Range("A15:A1000")
  For Each cell In MyPlage

                    If Left(cell.Value, 2) = "AA" And cell(1, 13).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
                    If Left(cell.Value, 2) = "AA" And cell(1, 13).Value <= 2 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If
        
              If Left(cell.Value, 2) = "AA" And cell(1, 13).Value >= 2 Then
        cell(1, 7).Interior.ColorIndex = 46
        End If
Next

Set MyPlage = Range("D15:D1000")
  For Each cell In MyPlage

                    If Left(cell.Value, 2) = "AA" And cell(1, 11).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
                    If Left(cell.Value, 2) = "AA" And cell(1, 11).Value <= 2 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If
        
        
                      If Left(cell.Value, 2) = "AA" And cell(1, 11).Value >= 2 Then
        cell(1, 4).Interior.ColorIndex = 46
        End If
Next
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
