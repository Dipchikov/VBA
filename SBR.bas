Attribute VB_Name = "SBR"
Sub SBR()
 '--------------------- XDB- 93---------------------------
    
    Set MyPlage = Range("D15:D1000")
 ActiveCell.FormulaR1C1 = _
        "=IF(ISBLANK(RC[-4]),""-"",INDEX(INDIRECT(R12C15),MATCH(RC[-10],'Standard length'!R1C1:R500C1,0),MATCH(RC[-7],'Standard length'!R1C1:R1C500,0)))+500"
    For Each cell In MyPlage
    
        If cell.Value = "XDB93" Then
            cell(1, 8).Value = ActiveCell.FormulaR1C1
        End If
        
        
    Next
    
End Sub
