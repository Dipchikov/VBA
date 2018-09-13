Attribute VB_Name = "clear_wiring_table"
Sub clear_wiring_table()


If ActiveSheet.Name = "Wiring table" Then
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.DisplayAlerts = False
 On Error Resume Next
ActiveSheet.ShowAllData
'
' delete Macro
'
answer = MsgBox("Are you sure you want to clear the table?" & vbNewLine & "Did you press the Routing/Sn:Err botton?", vbYesNo + vbQuestion, "Clear the table")
If answer = vbYes Then

Range("A15:N1000").Interior.ColorIndex = 0
Range("A15:L1000").ClearFormats
    Range("B1").Select
    Selection.ClearContents
    Range("O12").Select
    Selection.ClearContents
    Range("A15:L960").Select
    Selection.ClearContents
    Selection.Font.Size = 11
    Selection.Font.Name = "Calibri"
    Range("T15:T960").Select
    Selection.ClearContents
    Columns("C:C").Select
    Selection.NumberFormat = "General"
    Columns("F:F").Select
    Selection.NumberFormat = "General"
    Columns("E:E").Select
    Selection.NumberFormat = "@"
    Columns("B:B").Select
    Selection.NumberFormat = "@"

    
   Range("A15:L960").Select
    Range("L960").Activate
    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
    Selection.Font.Bold = False
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    

    '-------------Formulas---------------
    
    formula.formula
    
     '-------------Length formula---------------

    Range("K15:K960").formula = "=IF(ISBLANK(RC[-4]),""-"",INDEX(INDIRECT(R12C15),MATCH(RC[-10],'Standard length'!R1C1:R800C1,0),MATCH(RC[-7],'Standard length'!R1C1:R1C800,0)))"
    
     '-------------Cable type formula---------------
     
    Range("L15:L960").formula = "=IFNA(INDEX(INDIRECT(R12C13),MATCH(RC[-4],'Type of cables '!R2C1:R20C1,0),MATCH(RC[-5],'Type of cables '!R2C1:R2C20,0)),""-"")"
    
   '-------------Possible_errors---------------
    Possible_errors.Possible_errors


Application.Calculation = xlCalculationAutomatic
Application.DisplayAlerts = True
Application.ScreenUpdating = True
Range("A15").Select
    Else
    'do nothing
    End If
 End If

End Sub

