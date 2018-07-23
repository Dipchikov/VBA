Attribute VB_Name = "SaveAs"

Sub SaveAs()

'---------------------------------
If ActiveSheet.Name = "Wiring table" Then

   
    




    Dim lr As Long
    Dim InitialFoldr$
     ActiveWorkbook.Save
    
       If IsEmpty(Worksheets("Wiring table").Range("B1")) Then
        OutPut = MsgBox("Please add scheme number in cell B1!!!", vbOKOnly + vbExclamation)
        Exit Sub
        End If
    
     On Error Resume Next
     '-----------------scrips--------------------
    ActiveSheet.ShowAllData
    formula.formula
    '------------------CLEAR COLOUR FIRST -------------------------

    Range("A15:L1000").Interior.ColorIndex = 0
    '-----------------scrips--------------------
    Swap.Swap
    Legend_of_colours.Legend_of_colours
    soft_by_colour.soft_by_colour
    Routing.Routing
    CountColorValue.CountColorValue
    '-----------------Изтриване и копиране в WCT-------------------
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Sheets("WCT_form").Range("A15:L1048576").EntireRow.Delete
    Sheets("Wiring table").Select
    lr = Range("A" & Rows.Count).End(xlUp).Row
    Range("A1:l" & lr).Copy
    Sheets("WCT_form").Select
    Range("A1").Select
    ActiveSheet.Paste
    Range("A1").PasteSpecial Paste:=xlPasteValues
    Range("A1").PasteSpecial Paste:=xlPasteFormats
    Sheets("Wiring table").Select
    Range("A15").Select


    '---------Генериране на Нова страница------------------
    Dim wb As Workbook
    Set wb = Workbooks.Add
    Application.CopyObjectsWithCells = False
    ThisWorkbook.Sheets("WCT_form").Copy Before:=wb.Sheets(1)
    ActiveSheet.Name = Range("B1").Value
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
    
       '-------------Edit style---------------
    Columns("C:C").Select
    Selection.NumberFormat = "General"
    Columns("F:F").Select
    Selection.NumberFormat = "General"
        '-------------Formulas---------------
    Range("C15:C" & lr).formula = "=""-""&RC[-2]&"":""&RC[-1]"
   ' Range("C15").Select
    'ActiveCell.FormulaR1C1 = "=""-""&RC[-2]&"":""&RC[-1]"
   ' Range("C15").Select
    'Selection.AutoFill Destination:=Range("C15:C" & lr), Type:=xlFillDefault
    'Range("C15:C" & lr).Select
    Range("F15:F" & lr).formula = "=""-""&RC[-2]&"":""&RC[-1]"
    'Range("F15").Select
    'ActiveCell.FormulaR1C1 = "=""-""&RC[-2]&"":""&RC[-1]"
    'Range("F15").Select
   ' Selection.AutoFill Destination:=Range("F15:F" & lr), Type:=xlFillDefault
    'Range("F15:F" & lr).Select
    Range("A15").Select
    Application.CutCopyMode = False 'esp

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    

Dim sFileSaveName As Variant
Dim sPath As String

sPath = ActiveSheet.Range("B1").Value & "_WCT_reworked"
InitialFoldr$ = "\\10.28.38.5\ppmv\Productions\Italian\LVC\UniSec\!!!__Orders\!_____Ongoing Orders"
sFileSaveName = Application.GetSaveAsFilename(InitialFileName:=sPath, FileFilter:="Excel Files (*.xlsx), *.xlsm")
If sFileSaveName <> False Then
ActiveWorkbook.SaveAs sFileSaveName
End If
End If

End Sub

