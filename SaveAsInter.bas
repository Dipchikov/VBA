Attribute VB_Name = "SaveAsInter"
Sub SaveAsInter()

If ActiveSheet.Name = "Interconnections" Then


    Dim lr As Long
    Dim InitialFoldr$


       If IsEmpty(Worksheets("Interconnections").Range("B1")) Then
        OutPut = MsgBox("Please add scheme number in cell B1!!!", vbOKOnly + vbExclamation)
        Exit Sub
        End If
     If IsEmpty(Worksheets("Interconnections").Range("D1")) Then
      OutPut = MsgBox("Please add Project number in cell D1!!!", vbOKOnly + vbExclamation)
       Exit Sub
       End If
       
    ActiveWorkbook.Save
    


'------------------Изтриване на формата------------------------------------------

   Sheets("Interconnection_form").Range("A12:a1048576").EntireRow.Delete
       
     On Error Resume Next
    
   
    '---------wire colous------------
    wire_colours.wire_colours
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
    
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False


    Range("A1:j" & lr).Copy
    Sheets("Interconnection_form").Select
    Range("A1").Select
    ActiveSheet.Paste
    Range("A1").PasteSpecial Paste:=xlPasteValues
    Range("A1").PasteSpecial Paste:=xlPasteFormats

     Sheets("Interconnections").Activate
     Range("A12").Select

    'ActiveSheet.Name = Range("B1").Value



    Dim wb As Workbook
    Set wb = Workbooks.Add
    Application.CopyObjectsWithCells = False
    ThisWorkbook.Sheets("Interconnection_form").Copy Before:=wb.Sheets(1)
    ActiveSheet.Name = Range("D1").Value
    Application.CopyObjectsWithCells = True
    
    '---------Изтриване на Sheet1------------------
    Application.DisplayAlerts = False
    Sheets("Sheet1").Delete
    Application.DisplayAlerts = True
    


    '-------------add user in Footer ---------------
    With ActiveSheet.PageSetup
    .LeftFooter = "&D" & Chr(13) & "&9" & Application.UserName
    .RightFooter = "Page " & "&P" & Chr(13) & "&9" & Tools.Label8.Caption
    End With
    
         '-------------Formulas---------------
    Range("C12:C" & lr).formula = "=""-""&RC[-2]&"":""&RC[-1]"

    Range("F12:F" & lr).formula = "=""-""&RC[-2]&"":""&RC[-1]"

    Range("A6").Select
    
    
    Application.CutCopyMode = False 'esp
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    


Dim sFileSaveName As Variant
Dim sPath As String
sPath = "Interconnection_" & Right(ActiveSheet.Range("B1").Value, 4) & "_" & "Pos_" & ActiveSheet.Range("F1").Value
InitialFoldr$ = "\\10.28.38.5\ppmv\Productions\Italian\LVC\UniSec\!!!__Orders\!_____Ongoing Orders"
sFileSaveName = Application.GetSaveAsFilename(InitialFileName:=sPath, FileFilter:="Excel Files (*.xlsx), *.xlsm")
If sFileSaveName <> False Then
ActiveWorkbook.SaveAs sFileSaveName
Application.DisplayAlerts = False
ActiveWorkbook.Close
Application.DisplayAlerts = True
End If
End If
End Sub
