Attribute VB_Name = "Shielded_cable"
Sub Shielded_cable()
Dim lr As Long
lr = Range("A" & Rows.Count).End(xlUp).Row

Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False


'---------------------------XDI7----------------------------------------------
        
 Set MyPlage = Range("A15:A" & lr)
 For Each cell In MyPlage
        
If (cell.Value = "XDI7" Or cell(1, 4).Value = "XDI7") And Not (cell(1, 9).Value = "Saddle jumper" Or cell(1, 9).Value = "Insertable jumper") And cell(1, 12).Value <> "Shielded cable" Then
answer = MsgBox("Is connection between" & cell(1, 3).Value & " And " & cell(1, 6).Value & " is with Shielded cable", vbYesNo + vbQuestion, "-XDI7 Shielded cable")  'is with " & cell(1, 9).Value
        If answer = vbYes Then
        cell(1, 12).Value = "Shielded cable"
        'cell(1, 9).Font.ColorIndex = 3
        'cell(1, 9).Font.Bold = True
        End If
        End If
 
          
'If cell(1, 4).Value = "XDI7" And cell(1, 12).Value <> "Shielded cable" And Not (cell(1, 9).Value = "Saddle jumper" Or cell(1, 9).Value = "Insertable jumper") Then
'answer = MsgBox("Is connection between" & cell(1, 3).Value & " And " & cell(1, 6).Value & " is with Shielded cable", vbYesNo + vbQuestion, "-XDI7 Shielded cable") 'is with " & cell(1, 9).Value
'If answer = vbYes Then

        'cell(1, 12).Value = "Shielded cable"
        'cell(1, 9).Font.ColorIndex = 3
        'cell(1, 9).Font.Bold = True
 'End If
 'End If
                
Next

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub
