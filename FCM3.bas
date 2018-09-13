Attribute VB_Name = "FCM3"
Sub FCM3()
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
Dim FCMm As Single
FCMm = motor
        '---------------------------FCM3----------------------------------------------
  'FCMm = InputBox("Please add cross-section of conductors FCMm circuit" & vbNewLine & "Cross-section of conductors for FCMm circuit  by default is = 2,5", "Cross-Section for FCMm circuit", "2,5")
  Set MyPlage = Range("A15:A1000")
  For Each cell In MyPlage
  

            If cell.Value = "FCM3" And cell(1, 2).Value = 1 And cell(1, 7).Value <> FCMm Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = FCMm
        End If
                    
                    If cell.Value = "FCM3" And cell(1, 2).Value = 3 And cell(1, 7).Value <> FCMm Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = FCMm
        End If
                   
                    If cell.Value = "FCM3" And cell(1, 2).Value = 2 And cell(1, 7).Value <> FCMm Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = FCMm
        End If
                    
                    If cell.Value = "FCM3" And cell(1, 2).Value = 4 And cell(1, 7).Value <> FCMm Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = FCMm
        End If
                   
 Next
         
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub
