Attribute VB_Name = "formula"
Sub formula()
Columns("C:C").Select
    Selection.NumberFormat = "General"
    Columns("F:F").Select
    Selection.NumberFormat = "General"
    
        '-------------Formulas---------------
    
    Range("C15").Select
    ActiveCell.FormulaR1C1 = "=""-""&RC[-2]&"":""&RC[-1]"
    Range("C15").Select
    Selection.AutoFill Destination:=Range("C15:C951"), Type:=xlFillDefault
    Range("C15:C951").Select
    Range("F15").Select
    ActiveCell.FormulaR1C1 = "=""-""&RC[-2]&"":""&RC[-1]"
    Range("F15").Select
    Selection.AutoFill Destination:=Range("F15:F951"), Type:=xlFillDefault
    Range("F15:F951").Select
    
        Columns("C:C").Select
    Selection.NumberFormat = "@"
    Columns("F:F").Select
    Selection.NumberFormat = "@"
    
    Range("A15").Select
    
    
End Sub
