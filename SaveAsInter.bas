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


       
     On Error Resume Next
    
    ActiveWorkbook.Save
    Routing_inter.Routing_inter
   
    'Workbooks("CALCULATION OF CABLE LENGHTS_TEMPLATE - Italy Secondary.xlsm").Activate
    ActiveWorkbook.ActiveSheet.Select
    lr = Range("A" & Rows.Count).End(xlUp).Row

    Range("A1:j" & lr).Copy
    Workbooks.Open Filename:="C:\UniSec\Interconnection_form.xls", ReadOnly:=True
    Workbooks("Interconnection_form.xls").Activate
    Sheets("Interconnection").Select
    Range("A1").PasteSpecial Paste:=xlPasteFormats
    Range("A1").PasteSpecial Paste:=xlPasteValues
    
    ActiveSheet.Name = Range("B2").Value
    
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
sPath = "Interconnection_" & Right(ActiveSheet.Range("B1").Value, 4) & "_" & Left(ActiveSheet.Range("E1").Value, 2) & "k"
InitialFoldr$ = "\\10.28.38.5\ppmv\Productions\Italian\LVC\UniSec\!!!__Orders\!_____Ongoing Orders"
sFileSaveName = Application.GetSaveAsFilename(InitialFileName:=sPath, fileFilter:="Excel Files (*.xls), *.xlsm")
If sFileSaveName <> False Then
ActiveWorkbook.SaveAs sFileSaveName
End If
End If
End Sub
