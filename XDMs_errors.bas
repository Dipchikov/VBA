Attribute VB_Name = "XDMs_errors"
Sub XDMs_errors()
Dim lr As Long
Dim XDM As Double 'Single
Dim name As String

For i = 0 To 10
If i = 0 Then
name = "XDM"
Else
name = "XDM" & i
End If

Set MyPlage = Range("A15:d1000")
Application.ScreenUpdating = False
lr = Range("A" & Rows.Count).End(xlUp).Row


Set myCell = MyPlage.Find(What:=name, LookIn:=xlValues)
If Not myCell Is Nothing Then
XDM = InputBox("Please add cross-section of conductors" & vbNewLine & "If " & name & " is from  current circuit usually the cross-section of conductors is = 4mm" & vbNewLine & "If " & name & " is  from  voltage circuit usually the cross-section of conductors is = 1,5mm", "Cross-Section for " & name, Range("G" & myCell.Row).Value)
End If
Set MyPlage = Range("A15:A" & lr)
  For Each cell In MyPlage
 
              If cell.Value = name And cell(1, 7).Value <> XDM Then ' Not IsEmpty(cell(1, 7).Value) And
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = XDM
        End If

 Next

 
  Set MyPlage = Range("D15:D" & lr)
  For Each cell In MyPlage
       
            If cell.Value = name And cell(1, 4).Value <> XDM Then  'Not IsEmpty(cell(1, 4).Value) And
        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = XDM
        End If

Next

Next i

End Sub
