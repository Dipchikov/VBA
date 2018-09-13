Attribute VB_Name = "XDMs_errors"
Sub XDMs_errors()
Dim lr As Long
Dim XDM As Double 'Single


For i = 1 To 10
Set MyPlage = Range("A15:d1000")
Application.ScreenUpdating = False
lr = Range("A" & Rows.Count).End(xlUp).Row


Set myCell = MyPlage.Find(What:="XDM" & i, LookIn:=xlValues)
If Not myCell Is Nothing Then
XDM = InputBox("Please add cross-section of conductors" & vbNewLine & "If XDM" & i & " is from  current circuit usually the cross-section of conductors is = 4mm" & vbNewLine & "If XDM" & i & " is  from  voltage circuit usually the cross-section of conductors is = 1,5mm", "Cross-Section for XDM" & i, Range("G" & myCell.Row).Value)
End If
Set MyPlage = Range("A15:A" & lr)
  For Each cell In MyPlage
 
              If cell.Value = "XDM" & i And cell(1, 7).Value <> XDM Then ' Not IsEmpty(cell(1, 7).Value) And
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = XDM
        End If

 Next

 
  Set MyPlage = Range("D15:D" & lr)
  For Each cell In MyPlage
       
            If cell.Value = "XDM" & i And cell(1, 4).Value <> XDM Then  'Not IsEmpty(cell(1, 4).Value) And
        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = XDM
        End If
        
Next
Next i

End Sub
