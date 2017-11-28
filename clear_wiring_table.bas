Attribute VB_Name = "clear_wiring_table"
Sub clear_wiring_table()


If ActiveSheet.Name = "Wiring table" Then
Application.ScreenUpdating = False
Application.DisplayAlerts = False
 On Error Resume Next
ActiveSheet.ShowAllData
'
' delete Macro
'
answer = MsgBox("Are you sure you want to clear the table? Did you press the Routing botton?", vbYesNo + vbQuestion, "Clear the table")
If answer = vbYes Then

Range("A15:N1000").Interior.ColorIndex = 0

    Range("B1").Select
    Selection.ClearContents
        Range("O12").Select
    Selection.ClearContents
    Range("A15:L651").Select
    Selection.ClearContents
    Columns("C:C").Select
    Selection.NumberFormat = "General"
    Columns("F:F").Select
    Selection.NumberFormat = "General"
    Columns("E:E").Select
    Selection.NumberFormat = "@"
    Columns("B:B").Select
    Selection.NumberFormat = "@"

    
   Range("A15:L651").Select
    Range("L651").Activate
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
    

    '-------------Formulas---------------
    
    Range("C15").Select
    ActiveCell.FormulaR1C1 = "=""-""&RC[-2]&"":""&RC[-1]"
    Range("C15").Select
    Selection.AutoFill Destination:=Range("C15:C651"), Type:=xlFillDefault
    Range("C15:C651").Select
    Range("F15").Select
    ActiveCell.FormulaR1C1 = "=""-""&RC[-2]&"":""&RC[-1]"
    Range("F15").Select
    Selection.AutoFill Destination:=Range("F15:F651"), Type:=xlFillDefault
    Range("F15:F651").Select
    
        Columns("C:C").Select
    Selection.NumberFormat = "@"
    Columns("F:F").Select
    Selection.NumberFormat = "@"
       '-------------Length formula---------------
        Range("K15").Select
        
    ActiveCell.FormulaR1C1 = _
        "=IF(ISBLANK(RC[-4]),""-"",INDEX(INDIRECT(R12C15),MATCH(RC[-10],'Standard length'!R1C1:R500C1,0),MATCH(RC[-7],'Standard length'!R1C1:R1C500,0)))"
    Range("K15").Select
    Selection.AutoFill Destination:=Range("K15:K651")
    Range("K15:K651").Select
    
     '-------------Cable type formula---------------
     
        Range("L15").Select
    ActiveCell.FormulaR1C1 = _
        "=IFNA(INDEX(INDIRECT(R12C13),MATCH(RC[-4],'Type of cables '!R2C1:R15C1,0),MATCH(RC[-5],'Type of cables '!R2C1:R2C15,0)),""-"")"
    Range("L15").Select
    Selection.AutoFill Destination:=Range("L15:L651")
    Range("L15:L651").Select
    
   '-------------Possible_errors---------------
    Possible_errors.Possible_errors

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Range("A15").Select

    Else
    'do nothing
    End If
 End If

End Sub
