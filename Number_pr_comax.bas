Attribute VB_Name = "Number_pr_comax"
Sub Number()

Application.ScreenUpdating = False
Dim i As Long
Sheets("Komax").Select

For i = 2 To 981
If i <= 99 Then
Cells(i, "CO").Value = 1
End If
If i > 99 And i <= 197 Then
Cells(i, "CO").Value = 2
End If
If i > 197 And i <= 295 Then
Cells(i, "CO").Value = 3
End If
If i > 295 And i <= 393 Then
Cells(i, "CO").Value = 4
End If
If i > 393 And i <= 491 Then
Cells(i, "CO").Value = 5
End If
If i > 491 And i <= 589 Then
Cells(i, "CO").Value = 6
End If
If i > 589 And i <= 687 Then
Cells(i, "CO").Value = 7
End If
If i > 687 And i <= 785 Then
Cells(i, "CO").Value = 8
End If
If i > 785 And i <= 883 Then
Cells(i, "CO").Value = 9
End If
If i > 883 And i <= 981 Then
Cells(i, "CO").Value = 10
End If
Next i

Application.ScreenUpdating = True
End Sub

