Attribute VB_Name = "portfolio"
Sub portfolio()

    Dim xWs As Worksheet
    Dim lrdata As Long
    Dim lastRowSearch As Long
    Dim wbDataSheet As Worksheet
    Dim wbOtherSheet As Worksheet
    Dim orderNumber As String

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set wbDataSheet = Sheets("Portfolio")




    lrdata = wbDataSheet.Range("A" & Rows.Count).End(xlUp).Row
    Set myDataRange = wbDataSheet.Range("A7:A" & lrdata)
    wbDataSheet.Range("A7:r" & lrdata).Interior.ColorIndex = 0
    For Each cell In myDataRange
        orderNumber = cell.Value


        For Each xWs In ActiveWorkbook.Worksheets
            If Not (xWs.Name = "Portfolio") Then
                xWs.Activate
                Set wbOtherSheet = xWs
                Name = xWs.Name
                 lastRowSearch = wbOtherSheet.Range("A" & Rows.Count).End(xlUp).Row
                With wbOtherSheet.Range("A2:A" & lastRowSearch).Select
               
                    Set code = Selection.Find(What:=orderNumber, After:=ActiveCell, LookIn:=xlFormulas, _
                        LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                        MatchCase:=True, SearchFormat:=False)

                If Not code Is Nothing Then
                    Range(cell(1, 1), cell(1, 18)).Interior.ColorIndex = 0
                    Exit For
                Else
                    Range(cell(1, 1), cell(1, 18)).Interior.ColorIndex = 3
                End If
            End With
        End If

        Next
    Next cell

wbDataSheet.Activate
wbDataSheet.Range("A5").Select
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub












