Attribute VB_Name = "formula"
Sub formula()

On Error Resume Next
ActiveSheet.ShowAllData
Columns("C:C").Select
    Selection.NumberFormat = "General"
    Columns("F:F").Select
    Selection.NumberFormat = "General"
    
        '-------------Formulas---------------
    Range("C15:C960").formula = "=""-""&RC[-2]&"":""&RC[-1]"

    Range("F15:F960").formula = "=""-""&RC[-2]&"":""&RC[-1]"

    
        Columns("C:C").Select
    Selection.NumberFormat = "@"
    Columns("F:F").Select
    Selection.NumberFormat = "@"
    
Range("A15").Select
    
End Sub
