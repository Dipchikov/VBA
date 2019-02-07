Attribute VB_Name = "serialNumberErrors"
Sub serialNumberErrors()
    ' ----дефиниране-------------
    Dim my_FileName As Variant
    Dim shematic As String
    Dim project As String
    Dim err As Double
    Dim Connections As Double
    Dim Routing As Double
    Dim myData As Workbook
    Dim mySerr As Workbook
    Dim lrdata As Long
    Dim lrSerr As Long


Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
' ----'??????????? ?? ?????????-------------
Set myData = ActiveWorkbook
Set myDataSheet = myData.Sheets("Statistic")

If ActiveSheet.name = "Interconnections" Then
project = Range("B1").Value
End If
If ActiveSheet.name = "Wiring table" Then
project = Range("G1").Value
End If

lrdata = myDataSheet.Range("A" & Rows.Count).End(xlUp).Row
projectNumber = InputBox("Project Number :", "Project number", project)
my_FileName = Application.GetOpenFilename(FileFilter:="Excel Files,*.xl*;*.xm*")

    If my_FileName <> False Then

        Workbooks.Open FileName:=my_FileName
        Set mySerr = ActiveWorkbook
        Set mySerrSheet = Sheets("Register")
        lrSerr = Range("A" & Rows.Count).End(xlUp).Row

    End If

    Set myDataRange = myDataSheet.Range("B15:B" & lrdata)
    For Each cell In myDataRange
        If cell.Value = project Then
            shematic = cell(1, 2).Value
            err = cell(1, 7).Value
            Routing = RoutingFormula(cell(1, 8).Value)
            Connections = cell(1, 9).Value
            
     Set MyPlage = mySerrSheet.Range("E15:E" & lrSerr)
            For Each c In MyPlage
             If c.Value = shematic Then
                c(1, 12).Value = Connections
                c(1, 13).Value = err
                c(1, 15).Value = Routing
                End If
            Next

            End If
    Next cell

Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True

    Application.DisplayAlerts = False
    ActiveWorkbook.Save
    Application.DisplayAlerts = True
    'ActiveWorkbook.Close
    myData.Activate


End Sub


