Attribute VB_Name = "Routing"
Sub Routing()

Set MyPlage = Worksheets("Routing").Range("A15:A1000")

    For Each cell In MyPlage
        If IsEmpty(Worksheets("Wiring table").Range("B1")) Then
        rou = MsgBox("Please add scheme number in cell B1!!!", vbExclamation)
            Exit Sub
            End If
          If cell.Value = Worksheets("Wiring table").Range("B1") Then
        cell(1, 2).Value = Worksheets("Wiring table").Range("F10")
        cell(1, 5).Value = 1
        End If
        Next
 End Sub
    
