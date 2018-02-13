Attribute VB_Name = "SaveAsInter"
Sub SaveAsInter()

If ActiveSheet.Name = "Interconnections" Then


    Dim lr As Long
    Dim InitialFoldr$


       If IsEmpty(Worksheets("Interconnections").Range("B1")) Then
        OutPut = MsgBox("Please add scheme number in cell B1!!!", vbOKOnly + vbExclamation)
        Exit Sub
        End If
     If IsEmpty(Worksheets("Interconnections").Range("B2")) Then
      OutPut = MsgBox("Please add Project number in cell B2!!!", vbOKOnly + vbExclamation)
       Exit Sub
       End If


'------------------Изтриване на формата------------------------------------------

   Sheets("Interconnection_form").Range("A12:a1048576").EntireRow.Delete
       
     On Error Resume Next
    
    ActiveWorkbook.Save
    Routing_inter.Routing_inter
   

    ActiveWorkbook.ActiveSheet.Select
    lr = Range("A" & Rows.Count).End(xlUp).Row
    
    
    '------------------Филтър------------------------------------------
    ActiveWorkbook.Worksheets("Interconnections").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Interconnections").AutoFilter.Sort.SortFields.Add Key _
        :=Range("A12:A" & lr), SortOn:=xlSortOnValues, order:=xlAscending, _
        DataOption:=xlSortNormal
        
   With ActiveWorkbook.Worksheets("Interconnections").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = True
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    Range("A1:j" & lr).Copy
    'Workbooks.Open Filename:="C:\UniSec\Interconnection_form.xls", ReadOnly:=True
    'Workbooks("Interconnection_form.xls").Activate
    
    Sheets("Interconnection_form").Select
    Range("A1").Select
    ActiveSheet.Paste
    Range("A1").PasteSpecial Paste:=xlPasteValues
    Range("A1").PasteSpecial Paste:=xlPasteFormats

    Sheets("Interconnections").Select
    Range("A12").Select

    'ActiveSheet.Name = Range("B1").Value

    
    Dim wb As Workbook
    Set wb = Workbooks.Add
    Application.CopyObjectsWithCells = False
    ThisWorkbook.Sheets("Interconnection_form").Copy Before:=wb.Sheets(1)
    ActiveSheet.Name = Range("B2").Value
    Application.CopyObjectsWithCells = True
    
    '---------Изтриване на Sheet1------------------
    Application.DisplayAlerts = False
    Sheets("Sheet1").Delete
    Application.DisplayAlerts = True
    
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    

    '-------------add user in Footer ---------------
    With ActiveSheet.PageSetup
    .LeftFooter = "&D" & Chr(13) & Application.UserName
    End With
    
         '-------------Formulas---------------
    
    Range("C12").Select
    ActiveCell.FormulaR1C1 = "=""=""&RC[-2]&"":""&RC[-1]"
    Range("C12").Select
    Selection.AutoFill Destination:=Range("C12:C" & lr), Type:=xlFillDefault
    Range("C12:C" & lr).Select
    Range("F12").Select
    ActiveCell.FormulaR1C1 = "=""=""&RC[-2]&"":""&RC[-1]"
    Range("F12").Select
    Selection.AutoFill Destination:=Range("F12:F" & lr), Type:=xlFillDefault
    Range("F12:F" & lr).Select
    Range("A6").Select
    
    
    Application.CutCopyMode = False 'esp
    



Dim sFileSaveName As Variant
Dim sPath As String
sPath = "Interconnection_" & Right(ActiveSheet.Range("B1").Value, 4) & "_" & "Pos:" & ActiveSheet.Range("E1").Value
InitialFoldr$ = "\\10.28.38.5\ppmv\Productions\Italian\LVC\UniSec\!!!__Orders\!_____Ongoing Orders"
sFileSaveName = Application.GetSaveAsFilename(InitialFileName:=sPath, fileFilter:="Excel Files (*.xlsx), *.xlsm")
If sFileSaveName <> False Then
ActiveWorkbook.SaveAs sFileSaveName
End If
End If
End Sub
