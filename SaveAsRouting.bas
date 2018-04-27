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
    
        '---------Изтриване на Sheet1------------------
    Application.DisplayAlerts = False
    Sheets("Sheet1").Delete
    Application.DisplayAlerts = True
    
    
    '-------------add user in Footer ---------------
    With ActiveSheet.PageSetup
    .LeftFooter = "&D" & Chr(13) & "&9" & Application.UserName
    .RightFooter = "Page " & "&P" & Chr(13) & "&9" & Tools.Label8.Caption
    End With
    
    Application.CutCopyMode = False 'esp

Dim sFileSaveName As Variant
Dim sPath As String
sPath = "Routing_" & Right(ActiveSheet.Range("A4").Value, 4) & "_" & ActiveSheet.Range("B4").Value
InitialFoldr$ = "\\10.28.38.5\ppmv\Productions\Italian\LVC\UniSec\!!!__Orders\!_____Ongoing Orders"
sFileSaveName = Application.GetSaveAsFilename(InitialFileName:=sPath, FileFilter:="Excel Files (*.xlsx), *.xlsm")
If sFileSaveName <> False Then
Application.DisplayAlerts = False
ActiveWorkbook.SaveAs sFileSaveName
Application.DisplayAlerts = True
End If
End If
End Sub
