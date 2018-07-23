Attribute VB_Name = "Delete_new"
Sub Delete_new()
 Dim xWs As Worksheet
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

 answer = MsgBox("Are you sure you want to delete generated documents?", vbYesNo + vbQuestion, "Deleting generated documents")
If answer = vbYes Then

 For Each xWs In ActiveWorkbook.Worksheets
        If Not (xWs.Name = "Control Card" Or xWs.Name = "Register" Or xWs.Name = "Label" Or xWs.Name = "Repair" Or xWs.Name = "Routing by week" Or xWs.Name = "Operators Card" Or xWs.Name = "Repair") Then
            xWs.Delete
        End If
    Next
End If
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

