Attribute VB_Name = "Clear_Interconnections"

Sub Clear_Interconnections()

If ActiveSheet.Name = "Interconnections" Then
' delete Macro Interconnections
'
answer = MsgBox("Are you sure you want to clear the table?", vbYesNo + vbQuestion, "Clear the table")
If answer = vbYes Then
     Range("B1").Select
    Selection.ClearContents
         Range("B2").Select
    Selection.ClearContents
             Range("E1").Select
    Selection.ClearContents
 Range("A6:J515").Select
    Selection.ClearContents
    Range("C6").Select
    ActiveCell.FormulaR1C1 = "=""=""&RC[-2]&"":""&RC[-1]"
    Range("C6").Select
    Selection.AutoFill Destination:=Range("C6:C515"), Type:=xlFillDefault
    Range("C6:C515").Select
    Range("F6").Select
    ActiveCell.FormulaR1C1 = "=""=""&RC[-2]&"":""&RC[-1]"
    Range("F6").Select
    Selection.AutoFill Destination:=Range("F6:F515"), Type:=xlFillDefault
    Range("F6:F515").Select

    Range("I6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISBLANK(RC[-8]),""-"",(MID(RC[-5],2,2)-MID(RC[-8],2,2))+1)"
    Range("I6").Select
    Selection.AutoFill Destination:=Range("I6:I515"), Type:=xlFillDefault
    Range("I6:I515").Select
    Range("J6").Select
    ActiveCell.FormulaR1C1 = _
        "=IFNA(INDEX(INDIRECT(R3C12),MATCH(RC[-3],'Type of cables '!R2C1:R15C1,0),MATCH(RC[-2],'Type of cables '!R2C1:R2C15,0)),""-"")"
    Range("J6").Select
    Selection.AutoFill Destination:=Range("J6:J515"), Type:=xlFillDefault
    Range("J6:J515").Select
    Range("A6").Select
    
    Else
    'do nothing
    End If
    End If
End Sub

