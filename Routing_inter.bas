Attribute VB_Name = "Routing_inter"
Sub Routing_inter()

If ActiveSheet.Name = "Interconnections" Then
Set MyPlage = Worksheets("Routing").Range("A15:A1000")

    For Each cell In MyPlage
    If IsEmpty(Worksheets("Interconnections").Range("D1")) Then
        rou = MsgBox("Please add scheme number in cell D1!!!", vbExclamation)
            Exit Sub
          
         End If
        If cell.Value = Worksheets("Interconnections").Range("D1") Then
        cell(1, 2).Value = Worksheets("Interconnections").Range("D8")
        cell(1, 3).Value = 1
        End If
        Next
        End If
    End Sub
