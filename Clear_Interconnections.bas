Attribute VB_Name = "Clear_Interconnections"

Sub Clear_Interconnections()

If ActiveSheet.Name = "Interconnections" Then
' delete Macro Interconnections
'
answer = MsgBox("Are you sure you want to clear the table?", vbYesNo + vbQuestion, "Clear the table")
If answer = vbYes Then
     Range("B2").Select
     Selection.ClearContents
     Range("E1").Select
    Selection.ClearContents
     Range("A12:J515").Select
    Selection.ClearContents
    Range("C12").Select
    ActiveCell.FormulaR1C1 = "=""=""&RC[-2]&"":""&RC[-1]"
    Range("C12").Select
    Selection.AutoFill Destination:=Range("C12:C515"), Type:=xlFillDefault
    Range("C12:C515").Select
    Range("F12").Select
    ActiveCell.FormulaR1C1 = "=""=""&RC[-2]&"":""&RC[-1]"
    Range("F12").Select
    Selection.AutoFill Destination:=Range("F12:F515"), Type:=xlFillDefault
    Range("F12:F515").Select

    Range("I12").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISBLANK(RC[-8]),""-"",(MID(RC[-5],2,2)-MID(RC[-8],2,2))+1)"
    Range("I12").Select
    Selection.AutoFill Destination:=Range("I12:I515"), Type:=xlFillDefault
    Range("I12:I515").Select
    Range("J12").Select
    ActiveCell.FormulaR1C1 = _
        "=IFNA(INDEX(INDIRECT(R3C12),MATCH(RC[-3],'Type of cables '!R2C1:R15C1,0),MATCH(RC[-2],'Type of cables '!R2C1:R2C15,0)),""-"")"
    Range("J12").Select
    Selection.AutoFill Destination:=Range("J12:J515"), Type:=xlFillDefault
    Range("J12:J515").Select
    Range("A12").Select
    
    Else
    'do nothing
    End If
    End If
End Sub

