Attribute VB_Name = "sernumerr"
Sub sernumerr()
' ----дефиниране-------------
Dim my_FileName As Variant
Dim shematic As String
Dim err As Single
Dim connections As Single
Dim Routing As Single
Dim mydata As Workbook
Dim myserr As Workbook

' ----'присвояване на стойности-------------
Set mydata = ThisWorkbook
Sheets("Wiring table").Select
err = Range("H10").Value
connections = Range("L10").Value
Routing = Range("F10").Value * 0.1 + 1.34
shematic = Range("B1").Value
If shematic = "" Then
 rou = MsgBox("Please add scheme number in cell B1!!!", vbExclamation)
Exit Sub
End If


my_FileName = Application.GetOpenFilename(FileFilter:="Excel Files,*.xl*;*.xm*")

If my_FileName <> False Then

Workbooks.Open FileName:=my_FileName

Sheets("Register").Select

Set MyPlage = Range("E15:E1048576")
    For Each cell In MyPlage
        If cell.Value = shematic Then
        cell(1, 12).Value = connections
        cell(1, 13).Value = err
        cell(1, 15).Value = Routing
        End If
        Next
        Else
        Exit Sub
        End If
    Application.DisplayAlerts = False
    ActiveWorkbook.Save
    Application.DisplayAlerts = True
    'ActiveWorkbook.Close
    mydata.Activate

End Sub
