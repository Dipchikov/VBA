Attribute VB_Name = "SaveAs"

Sub SaveAs()

'---------------------------------
If ActiveSheet.Name = "Wiring table" Then

Routing_inter.Routing_inter

    Dim lr As Long
    Dim InitialFoldr$
     ActiveWorkbook.Save
    
       If IsEmpty(Worksheets("Wiring table").Range("B1")) Then
        MsgBox "Please add scheme number in cell B1!!!"
        Exit Sub
        End If
    
     On Error Resume Next
     '-----------------scrips--------------------
    ActiveSheet.ShowAllData
    Swap.Swap
    Legend_of_colours.Legend_of_colours
    soft_by_colour.soft_by_colour
    Routing.Routing




   
    'Workbooks("CALCULATION OF CABLE LENGHTS_TEMPLATE - Italy Secondary.xlsm").Activate
    Sheets("Wiring table").Select
    lr = Range("A" & Rows.Count).End(xlUp).Row

    Range("A1:l" & lr).Copy
    Workbooks.Open Filename:="C:\UniSec\CONNECTION_LIST_form.xls", ReadOnly:=True
    Workbooks("CONNECTION_LIST_form.xls").Activate
    Sheets("LISTA CONNESSIONI1").Select
    Range("A1").Select
    ActiveSheet.Paste
    Range("A1").PasteSpecial Paste:=xlPasteValues
    'Range("A1").PasteSpecial Paste:=xlPasteFormats
    
    ActiveSheet.Name = Range("B1").Value
       '-------------Edit style---------------
    Columns("C:C").Select
    Selection.NumberFormat = "General"
    Columns("F:F").Select
    Selection.NumberFormat = "General"
        '-------------Formulas---------------
    Range("C15").Select
    ActiveCell.FormulaR1C1 = "=""-""&RC[-2]&"":""&RC[-1]"
    Range("C15").Select
    Selection.AutoFill Destination:=Range("C15:C" & lr), Type:=xlFillDefault
    Range("C15:C" & lr).Select
    Range("F15").Select
    ActiveCell.FormulaR1C1 = "=""-""&RC[-2]&"":""&RC[-1]"
    Range("F15").Select
    Selection.AutoFill Destination:=Range("F15:F" & lr), Type:=xlFillDefault
    Range("F15:F" & lr).Select
    Range("A15").Select

    Application.CutCopyMode = False 'esp


Dim sFileSaveName As Variant
Dim sPath As String


sPath = Workbooks("CONNECTION_LIST_form.xls").ActiveSheet.Range("B1").Value & "_CONNECTION_LIST_reworked"
InitialFoldr$ = "\\10.28.38.5\ppmv\Productions\Italian\LVC\UniSec\!!!__Orders\!_____Ongoing Orders"
sFileSaveName = Application.GetSaveAsFilename(InitialFileName:=sPath, fileFilter:="Excel Files (*.xls), *.xlsm")
If sFileSaveName <> False Then
ActiveWorkbook.SaveAs sFileSaveName
End If
End If
End Sub
