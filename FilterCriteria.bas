Attribute VB_Name = "FilterCriteria"
Sub FilterCriteria()
    Dim filterValues() As Variant, cl As Range, i As Integer

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    If Not IsEmpty(ActiveCell.Value) Then
    ReDim filterValues(Selection.Cells.Count - 1)
    i = 0
    For Each cl In Selection
    
        filterValues(i) = cl.text
        i = i + 1
    Next cl
    Range(ActiveCell.CurrentRegion.Address).AutoFilter Field:=ActiveCell.Column, Criteria1:=filterValues, Operator:=xlFilterValues
    
    Else
    MsgBox "Please select not empty cell!"
    End If
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
