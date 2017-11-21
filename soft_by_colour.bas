Attribute VB_Name = "soft_by_colour"
Sub soft_by_colour()
Attribute soft_by_colour.VB_ProcData.VB_Invoke_Func = " \n14"

Application.ScreenUpdating = False
' soft_by_colour Macro
'----------------start Legend_of_colours-----------------
'Legend_of_colours.Legend_of_colours

    lr = Range("A" & Rows.Count).End(xlUp).Row

    ActiveWorkbook.Worksheets("Wiring table").AutoFilter.Sort.SortFields.Clear
    '--------------------Refs-----------------------------------------------------
    ActiveWorkbook.Worksheets("Wiring table").AutoFilter.Sort.SortFields.Add(Range( _
        "K15:K" & lr), xlSortOnCellColor, xlAscending, , xlSortNormal).SortOnValue.Color _
        = RGB(255, 204, 0)
    '--------------------Doors-----------------------------------------------------
        ActiveWorkbook.Worksheets("Wiring table").AutoFilter.Sort.SortFields.Add(Range( _
        "K15:K" & lr), xlSortOnCellColor, xlAscending, , xlSortNormal).SortOnValue.Color _
        = RGB(153, 204, 0)
        
     '--------------------Inside-----------------------------------------------------
    ActiveWorkbook.Worksheets("Wiring table").AutoFilter.Sort.SortFields.Add(Range( _
        "K15:K" & lr), xlSortOnCellColor, xlAscending, , xlSortNormal).SortOnValue.Color _
        = RGB(255, 204, 153)
  
  '--------------------Shielded cable----------------------------------------------------
    ActiveWorkbook.Worksheets("Wiring table").AutoFilter.Sort.SortFields.Add(Range( _
        "K15:K" & lr), xlSortOnCellColor, xlAscending, , xlSortNormal).SortOnValue.Color _
        = RGB(255, 255, 0)
               
               
        '--------------------XDB-----------------------------------------------------
    ActiveWorkbook.Worksheets("Wiring table").AutoFilter.Sort.SortFields.Add(Range( _
        "K15:K" & lr), xlSortOnCellColor, xlAscending, , xlSortNormal).SortOnValue.Color _
        = RGB(153, 204, 255)
        '--------------------Jumpers-----------------------------------------------------
    ActiveWorkbook.Worksheets("Wiring table").AutoFilter.Sort.SortFields.Add(Range( _
        "K15:K" & lr), xlSortOnCellColor, xlAscending, , xlSortNormal).SortOnValue.Color _
        = RGB(128, 128, 128)
    With ActiveWorkbook.Worksheets("Wiring table").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    Application.ScreenUpdating = True

End Sub
