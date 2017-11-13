Attribute VB_Name = "Routing_inter"
Sub Routing_inter()
If ActiveSheet.Name = "Interconnections" Then
Set MyPlage = Worksheets("Routing").Range("A15:A1000")

    For Each cell In MyPlage
    If IsEmpty(Worksheets("Interconnections").Range("B2")) Then
        MsgBox "Please add scheme number in cell B2!!!"
            Exit Sub
          
            End If
          If cell.Value = Worksheets("Interconnections").Range("B2") Then
        cell(1, 2).Value = Worksheets("Interconnections").Range("J1")
        cell(1, 3).Value = 1
        End If
        Next
        End If
    End Sub
