Attribute VB_Name = "Possible_errors"
Sub Possible_errors()
Attribute Possible_errors.VB_ProcData.VB_Invoke_Func = " \n14"

'
'  Possible_errors
'

    Range("M15").Select
    ActiveCell.FormulaR1C1 = "=IF(RC3=""-:"",,COUNTIF(R15C3:R960C6,RC[-10]))"
    Range("M15").Select
    Selection.AutoFill Destination:=Range("M15:M960")
    Range("M15:M960").Select
    Range("N15").Select
    ActiveCell.FormulaR1C1 = "=IF(RC6=""-:"",,COUNTIF(R15C3:R960C6,RC[-8]))"
    Range("N15").Select
    Selection.AutoFill Destination:=Range("N15:N960")
    Range("N15:N960").Select
    
    
    
End Sub

