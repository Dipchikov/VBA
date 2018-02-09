Attribute VB_Name = "SaveAsRouting"
Sub SaveAsRouting()
If ActiveSheet.Name = "Routing" Then
    Dim lr As Long
    Dim InitialFoldr$
    
    ActiveWorkbook.Save
    
   
    'Workbooks("CALCULATION OF CABLE LENGHTS_TEMPLATE - Italy Secondary.xlsm").Activate
    ActiveWorkbook.ActiveSheet.Select
    lr = Range("A" & Rows.Count).End(xlUp).Row
    
    
    Dim wb As Workbook
    Set wb = Workbooks.Add
    Application.CopyObjectsWithCells = False
    ThisWorkbook.Sheets("Routing").Copy Before:=wb.Sheets(1)
    Application.CopyObjectsWithCells = True
    'Range("A2:K" & lr).Copy
    'Workbooks.Open Filename:="C:\UniSec\Routing_form.xlsx", ReadOnly:=True
   ' Workbooks("Routing_form.xlsx").Activate
   ' Sheets("UNISEC").Select
   '    ' Range("A2").Select
    'ActiveSheet.Paste
   ' Range("A2").PasteSpecial Paste:=xlPasteFormats

         
  
    Application.CutCopyMode = False 'esp

Dim sFileSaveName As Variant
Dim sPath As String
sPath = "Routing_" & Right(ActiveSheet.Range("A4").Value, 4) & "_" & Left(ActiveSheet.Range("B4").Value, 2) & "k"
InitialFoldr$ = "\\10.28.38.5\ppmv\Productions\Italian\LVC\UniSec\!!!__Orders\!_____Ongoing Orders"
sFileSaveName = Application.GetSaveAsFilename(InitialFileName:=sPath, fileFilter:="Excel Files (*.xlsx), *.xlsm")
If sFileSaveName <> False Then
ActiveWorkbook.SaveAs sFileSaveName
End If
End If
End Sub
