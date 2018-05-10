Attribute VB_Name = "XDMs_errors"
Sub XDMs_errors()

Dim XDM1 As String

Set MyPlage = Range("A15:d1000")


Set myCell = MyPlage.Find(What:="XDM1", LookIn:=xlValues)
If Not myCell Is Nothing Then
XDM1 = InputBox("Please add cross-section of conductors" & vbNewLine & "If XDM1 is from  current circuit the cross-section of conductors need to be = 4mm" & vbNewLine & "If XDM1 is  from  voltage circuit the cross-section of conductors need to be = 1,5mm", "Cross-Section for XDM1", 4)
End If
Set MyPlage = Range("A15:A1000")
  For Each cell In MyPlage
 
              If Not IsEmpty(cell(1, 7).Value) And cell.Value = "XDM1" And cell(1, 7).Value <> XDM1 Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = XDM1
        End If
  
 Next
 
 
    Set MyPlage = Range("D15:D1000")
  For Each cell In MyPlage
       
            If Not IsEmpty(cell(1, 4).Value) And cell.Value = "XDM1" And cell(1, 4).Value <> XDM1 Then
        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = XDM1
        End If
        
Next

 Dim XDM2 As String
Set MyPlage = Range("A15:d1000")


Set myCell = MyPlage.Find(What:="XDM2", LookIn:=xlValues)
If Not myCell Is Nothing Then
XDM2 = InputBox("Please add cross-section of conductors" & vbNewLine & "If XDM2 is from  current circuit the cross-section of conductors need to be = 4mm" & vbNewLine & "If XDM2 is  from  voltage circuit the cross-section of conductors need to be = 1,5mm", "Cross-Section for XDM2", "1,5")
End If
Set MyPlage = Range("A15:A1000")
  For Each cell In MyPlage
 
              If Not IsEmpty(cell(1, 7).Value) And cell.Value = "XDM2" And cell(1, 7).Value <> XDM2 Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = XDM2
        End If
  
 Next
 
 
    Set MyPlage = Range("D15:D1000")
  For Each cell In MyPlage
       
            If Not IsEmpty(cell(1, 4).Value) And cell.Value = "XDM2" And cell(1, 4).Value <> XDM2 Then
        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = XDM2
        End If
        
Next

Dim XDM3 As String
Set MyPlage = Range("A15:d1000")


Set myCell = MyPlage.Find(What:="XDM3", LookIn:=xlValues)
If Not myCell Is Nothing Then
XDM3 = InputBox("Please add cross-section of conductors" & vbNewLine & "If XDM3 is from  current circuit the cross-section of conductors need to be = 4mm" & vbNewLine & "If XDM3 is  from  voltage circuit the cross-section of conductors need to be = 1,5mm", "Cross-Section for XDM3", "1,5")
End If
Set MyPlage = Range("A15:A1000")
  For Each cell In MyPlage
 
              If Not IsEmpty(cell(1, 7).Value) And cell.Value = "XDM3" And cell(1, 7).Value <> XDM3 Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = XDM3
        End If
  
 Next
 
 
    Set MyPlage = Range("D15:D1000")
  For Each cell In MyPlage
       
            If Not IsEmpty(cell(1, 4).Value) And cell.Value = "XDM3" And cell(1, 4).Value <> XDM3 Then
        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = XDM3
        End If
        
Next
   
Dim XDM4 As String
Set MyPlage = Range("A15:d1000")


Set myCell = MyPlage.Find(What:="XDM4", LookIn:=xlValues)
If Not myCell Is Nothing Then
XDM4 = InputBox("Please add cross-section of conductors" & vbNewLine & "If XDM4 is from  current circuit the cross-section of conductors need to be = 4mm" & vbNewLine & "If XDM4 is  from  voltage circuit the cross-section of conductors need to be = 1,5mm", "Cross-Section for XDM4", "1,5")
End If
Set MyPlage = Range("A15:A1000")
  For Each cell In MyPlage
 
              If Not IsEmpty(cell(1, 7).Value) And cell.Value = "XDM4" And cell(1, 7).Value <> XDM4 Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = XDM4
        End If
  
 Next
 
 
    Set MyPlage = Range("D15:D1000")
  For Each cell In MyPlage
       
            If Not IsEmpty(cell(1, 4).Value) And cell.Value = "XDM4" And cell(1, 4).Value <> XDM4 Then
        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = XDM4
        End If
        
Next

Dim XDM5 As String
Set MyPlage = Range("A15:d1000")


Set myCell = MyPlage.Find(What:="XDM5", LookIn:=xlValues)
If Not myCell Is Nothing Then
XDM5 = InputBox("Please add cross-section of conductors" & vbNewLine & "If XDM5 is from  current circuit the cross-section of conductors need to be = 4mm" & vbNewLine & "If XDM5 is  from  voltage circuit the cross-section of conductors need to be = 1,5mm", "Cross-Section for XDM5", "1,5")
End If
Set MyPlage = Range("A15:A1000")
  For Each cell In MyPlage
 
              If Not IsEmpty(cell(1, 7).Value) And cell.Value = "XDM5" And cell(1, 7).Value <> XDM5 Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = XDM5
        End If
  
 Next
 
 
    Set MyPlage = Range("D15:D1000")
  For Each cell In MyPlage
       
            If Not IsEmpty(cell(1, 4).Value) And cell.Value = "XDM5" And cell(1, 4).Value <> XDM5 Then
        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = XDM5
        End If
        
Next

End Sub
