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
    Selection.AutoFill Destination:=Range("CO2:CO99"), Type:=xlFillDefault
    Range("CO2:CO99").Select
    Range("CO100").Select
    ActiveCell.FormulaR1C1 = "2"
    Range("CO100").Select
    Selection.AutoFill Destination:=Range("CO100:CO195"), Type:=xlFillDefault
    Range("CO100:CO195").Select
    Range("CO196").Select
    ActiveCell.FormulaR1C1 = "3"
    Range("CO196").Select
    Selection.AutoFill Destination:=Range("CO196:CO291"), Type:=xlFillDefault
    Range("CO196:CO291").Select
     Range("CO292").Select
    ActiveCell.FormulaR1C1 = "4"
    Range("CO292").Select
    Selection.AutoFill Destination:=Range("CO292:CO387"), Type:=xlFillDefault
    Range("CO292:CO387").Select
    Range("CO388").Select
    ActiveCell.FormulaR1C1 = "5"
    Range("CO388").Select
    Selection.AutoFill Destination:=Range("CO388:CO483"), Type:=xlFillDefault
    Range("CO388:CO483").Select
    Range("CO484").Select
    ActiveCell.FormulaR1C1 = "6"
    Range("CO484").Select
    Selection.AutoFill Destination:=Range("CO484:CO579"), Type:=xlFillDefault
    Range("CO484:CO579").Select
    Range("CO580").Select
    ActiveCell.FormulaR1C1 = "7"
    Range("CO580").Select
    Selection.AutoFill Destination:=Range("CO580:CO675"), Type:=xlFillDefault
	Range("CO676").Select
    ActiveCell.FormulaR1C1 = "8"
    Range("CO676").Select
    Selection.AutoFill Destination:=Range("CO676:CO771"), Type:=xlFillDefault
    Range("A2").Select

End Sub
