Attribute VB_Name = "FCM3"
Sub FCM3()
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
        '---------------------------FCM3----------------------------------------------
        
        Set MyPlage = Range("A15:A1000")
  For Each cell In MyPlage
  

            If cell.Value = "FCM3" And cell(1, 2).Value = 1 And cell(1, 7).Value < "2,5" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "2,5"
        End If
                    
                    If cell.Value = "FCM3" And cell(1, 2).Value = 3 And cell(1, 7).Value < "2,5" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "2,5"
        End If
                   
                    If cell.Value = "FCM3" And cell(1, 2).Value = 2 And cell(1, 7).Value < "2,5" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "2,5"
        End If
                    
                    If cell.Value = "FCM3" And cell(1, 2).Value = 4 And cell(1, 7).Value < "2,5" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "2,5"
        End If
                   
 Next
         
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub
