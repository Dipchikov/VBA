Attribute VB_Name = "Number_pr_comax"
Sub number()
Attribute number.VB_ProcData.VB_Invoke_Func = " \n14"
'
' number Macro
'
Sheets("Comax").Select
    Range("CO2").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("CO2").Select
    Selection.AutoFill Destination:=Range("CO2:CO80"), Type:=xlFillDefault
    Range("CO2:CO80").Select
    Range("CO81").Select
    ActiveCell.FormulaR1C1 = "2"
    Range("CO81").Select
    Selection.AutoFill Destination:=Range("CO81:CO161"), Type:=xlFillDefault
    Range("CO81:CO161").Select
    Range("CO162").Select
    ActiveCell.FormulaR1C1 = "3"
    Range("CO162").Select
    Selection.AutoFill Destination:=Range("CO162:CO241"), Type:=xlFillDefault
    Range("CO162:CO241").Select
    ActiveWindow.SmallScroll Down:=12
    Range("CO242").Select
    ActiveCell.FormulaR1C1 = "4"
    Range("CO242").Select
    Selection.AutoFill Destination:=Range("CO242:CO322"), Type:=xlFillDefault
    Range("CO242:CO322").Select
    ActiveWindow.SmallScroll Down:=9
    Range("CO323").Select
    ActiveCell.FormulaR1C1 = "5"
    Range("CO323").Select
    Selection.AutoFill Destination:=Range("CO323:CO403"), Type:=xlFillDefault
    Range("CO323:CO403").Select
    Range("CO404").Select
    ActiveCell.FormulaR1C1 = "6"
    Range("CO404").Select
    Selection.AutoFill Destination:=Range("CO404:CO484"), Type:=xlFillDefault
    Range("CO404:CO484").Select
    ActiveWindow.SmallScroll Down:=9
    Range("CO485").Select
    ActiveCell.FormulaR1C1 = "7"
    Range("CO485").Select
    Selection.AutoFill Destination:=Range("CO485:CO565"), Type:=xlFillDefault
    Range("CO485:CO565").Select

End Sub
