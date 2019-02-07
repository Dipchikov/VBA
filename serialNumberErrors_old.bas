Attribute VB_Name = "serialNumberErrors_old"
Sub serialNumberErrors()
' ----дефиниране-------------
Dim my_FileName As Variant
Dim shematic As String
Dim err As Double
Dim Connections As Double
Dim Routing As Double
Dim myData As Workbook
Dim mySerr As Workbook
Dim lr As Long

' ----'??????????? ?? ?????????-------------
Set myData = ThisWorkbook
Sheets("Wiring table").Select
err = Range("H10").Value
Connections = Range("L10").Value
Routing = RoutingFormula(Range("F10").Value)
shematic = Range("B1").Value
If shematic = "" Then
 rou = MsgBox("Please add scheme number in cell B1!!!", vbExclamation)
Exit Sub
End If


my_FileName = Application.GetOpenFilename(FileFilter:="Excel Files,*.xl*;*.xm*")

If my_FileName <> False Then

Workbooks.Open FileName:=my_FileName

Sheets("Register").Select

lr = Range("E" & Rows.Count).End(xlUp).Row
Set MyPlage = Range("E15:E" & lr)
    For Each cell In MyPlage
        If cell.Value = shematic Then
        cell(1, 12).Value = Connections
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
    myData.Activate

End Sub

