Attribute VB_Name = "Komax_table_inter"
Sub Komax_table_inter()

 On Error Resume Next
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
If ActiveSheet.Name = "Interconnections" Then
       If IsEmpty(Worksheets("Interconnections").Range("B1")) Then
        OutPut = MsgBox("Please add scheme number in cell B1!!!", vbOKOnly + vbExclamation)
        Exit Sub
        End If
     If IsEmpty(Worksheets("Interconnections").Range("B2")) Then
      OutPut = MsgBox("Please add Project number in cell B2!!!", vbOKOnly + vbExclamation)
       Exit Sub
       End If

Set Data = Sheets("Interconnections")
Set Final = Sheets("Komax")
Data.Active
Data.ShowAllData
'Swap.Swap
'soft_by_colour.soft_by_colour



      '--------------------------- Clar teble-----------------

    Final.Range("A2").Select
    Final.Range("A2:CO1000").ClearContents
      '----------Prigram number------------------
    Number_pr_comax.number
    Final.Range("A2").Select
    
Dim i As Long
    Set Rng = Data.Range("J12:J1048576")
    For i = Rng.Cells(1, 1).Row To Rng.Cells(1, 1).End(xlDown).Row
        '----------Condition If cell is empty-------------------
        If Not (Data.Range("J" & i).Value = "-" Or Data.Range("J" & i).Value = "Shielded cable") Then
            'Final.Range("A" & i - 10).Value = "INTERP" & Left(Data.Range("E1").Value, 2) & "." & Final.Range("CO" & i - 10).Value
            Final.Range("A" & i - 10).Value = Data.Range("B2").Value & "." & Final.Range("CO" & i - 10).Value
            Final.Range("C" & i - 10).Value = 1
            Final.Range("D" & i - 10).Value = 1
            Final.Range("E" & i - 10).Value = "WA for " & Data.Range("B2").Value
            Final.Range("G" & i - 10).Value = Final.Range("A" & i - 10).Value
            Final.Range("I" & i - 10).Value = Final.Range("E" & i - 10).Value
            Final.Range("H" & i - 10).Value = "Italy\UniSec\" & Right(Data.Range("B1").Value, 4) & "####"
            Final.Range("M" & i - 10).Value = Data.Range("I" & i).Value * 1000
            Final.Range("K" & i - 10).Value = Data.Range("J" & i).Value
            Final.Range("AG" & i - 10).Value = "'" & Data.Range("C" & i).Value
            Final.Range("AH" & i - 10).Value = "'" & Data.Range("C" & i).Value
            Final.Range("AI" & i - 10).Value = "'" & Data.Range("C" & i).Value
            Final.Range("AJ" & i - 10).Value = 0
            Final.Range("AK" & i - 10).Value = "'" & Data.Range("F" & i).Value
            Final.Range("AL" & i - 10).Value = "'" & Data.Range("F" & i).Value
            Final.Range("AM" & i - 10).Value = "'" & Data.Range("F" & i).Value
            Final.Range("AO" & i - 10).Value = "'" & Data.Range("F" & i).Value
            Final.Range("AN" & i - 10).Value = 1
            Final.Range("AP" & i - 10).Value = 1
            Final.Range("BC" & i - 10).Value = 1
        '----------Cut for ferules-- StrippingLength-----------------
        If Not Data.Range("H" & i).Value = 4 Then
            Final.Range("O" & i - 10).Value = 10
            Final.Range("P" & i - 10).Value = 10
            Final.Range("BA" & i - 10).Value = 1
            Else
            Final.Range("O" & i - 10).Value = 12
            Final.Range("P" & i - 10).Value = 12
            Final.Range("BA" & i - 10).Value = 1
        End If
        
        End If
    Next i
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    Sheets("Interconnections").Select
    Range("A6").Select


 Final.Range("A:A").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    lr = Final.Range("A" & Rows.Count).End(xlUp).Row
    
        Final.Range("CO1").Select
    Final.Range("CO1:CO1048576").ClearContents
    
    Dim wb As Workbook
    Set wb = Workbooks.Add
    ThisWorkbook.Sheets("Komax").Copy Before:=wb.Sheets(1)
    
 'Final.Range("A1:CB" & lr).Copy
  'Workbooks.Open Filename:="C:\UniSec\Comax_form.csv", ReadOnly:=True
   ' Workbooks("Comax_form.csv").Activate
    'Range("A1").Select
    'ActiveSheet.Paste
   ' Range("A1").PasteSpecial Paste:=xlPasteValues

Application.CutCopyMode = False 'esp

Dim sFileSaveName As Variant
Dim sPath As String
'sPath = Workbooks("Comax_form.csv").ActiveSheet.Range("A2").Value
sPath = "Inter_" & Right(Data.Range("B1").Value, 4) & "_" & ActiveSheet.Range("A2").Value
InitialFoldr$ = "\\10.28.38.5\ppmv\Productions\Italian\LVC\UniSec\!!!__Orders\!_____Ongoing Orders"
sFileSaveName = Application.GetSaveAsFilename(InitialFileName:=Left(sPath, 26), fileFilter:="Excel Files (*.csv), *.xlsm")
If sFileSaveName <> False Then
ActiveWorkbook.SaveAs sFileSaveName, FileFormat:=xlCSV, Local:=True

End If
Else
 answer = MsgBox("To generate Comax table please make Worksheet Interconnections active !!!", vbYes + vbQuestion, "")
End If

End Sub


