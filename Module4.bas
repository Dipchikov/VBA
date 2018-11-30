Attribute VB_Name = "Module4"
Sub test()

lr = Range("A" & Rows.Count).End(xlUp).Row
 '------------------Jumpers between equipment -------------------------
Dim x As Double
Set MyPlage = Range("A15:A" & lr)

    For Each cell In MyPlage
        
        If cell.Value <> cell(1, 4).Value Then
         If cell(1, 9).Value <> "Direct Connection" Then
         If cell(1, 9).Value <> "Conductor / wire" Then
            cell(1, 9).Value = "Conductor / wire"
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
            End If
       
            If IsEmpty(cell(1, 7).Value) Then
             x = InputBox("Please add cross-section of conductors between" & vbNewLine & cell(1, 3) & " and " & cell(1, 6), "Wire jumper  between " & cell(1, 3) & " and " & cell(1, 6), wire, 1)
            cell(1, 7).Value = x
            cell(1, 7).Font.ColorIndex = 3
            cell(1, 7).Font.Bold = True
        End If
        If IsEmpty(cell(1, 8).Value) Then
        cell(1, 8).Value = "bk"
        cell(1, 8).Font.ColorIndex = 3
        cell(1, 8).Font.Bold = True
        End If
        End If
         End If
 Next
        
End Sub
