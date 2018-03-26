Attribute VB_Name = "Komax_table"
Sub Komax_table()

 On Error Resume Next
If ActiveSheet.Name = "Wiring table" Then

       If IsEmpty(Worksheets("Wiring table").Range("B1")) Then
        OutPut = MsgBox("Please add scheme number in cell B1!!!", vbOKOnly + vbExclamation)
        Exit Sub
        End If
     If IsEmpty(Worksheets("Wiring table").Range("G1")) Then
      OutPut = MsgBox("Please add Project number in cell G1!!!", vbOKOnly + vbExclamation)
       Exit Sub
        End If


formula.formula

Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
Set Data = Sheets("Wiring table")
Set Final = Sheets("Komax")
Data.Active
Data.ShowAllData
Swap.Swap

'Legend_of_colours.Legend_of_colours
'soft_by_colour.soft_by_colour

    '-----------------------Дефиниране на първия ред с "-"-----------------
    Dim lRow As Long
    Set Rng1 = Data.Range("L15:L1048576")
    lRow = Rng1.Find(What:="-", LookIn:=xlValues).Row
    
    '-----------------------Legend_of_feruless-----------------
    answer = MsgBox("Is this have a Weidmuller terminals ?" & vbNewLine & "And if this project have then press - Yes", vbYesNo + vbQuestion + vbDefaultButton2, "Comax table")
    If answer = vbYes Then
    UserForm1.Show vbModal
    Else
    Legend_of_feruless.Legend_of_feruless
    End If
      '--------------------------- Clar teble-----------------

    Final.Range("A2").Select
    Final.Range("A2:CO1048576").EntireRow.Delete
        '----------Prigram number------------------
    Number_pr_comax.Number
    
    Final.Range("A2").Select
    Set Rng = Data.Range("L15:L1048576")
    For i = Rng.Cells(1, 1).Row To Rng.Cells(1, 1).End(xlDown).Row

        '----------Condition If cell is empty------------------
        
        If Not (Data.Range("L" & i).Value = "-" Or Data.Range("L" & i).Value = "Shielded cable") Then
        '-----------------------Дефиниране на праграма под 99 реда-----------------
            If lRow <= 113 Then
            Final.Range("A" & i - 13).Value = Left(Data.Range("B1").Value, 10) & "W" & Right(Data.Range("B1").Value, 4)
            Else
            Final.Range("A" & i - 13).Value = Left(Data.Range("B1").Value, 10) & "W" & Right(Data.Range("B1").Value, 4) & "." & Final.Range("CO" & i - 13).Value
            End If 'край на условието
            Final.Range("C" & i - 13).Value = 1
            Final.Range("D" & i - 13).Value = 1
            Final.Range("E" & i - 13).Value = "WA for " & Data.Range("B1").Value
            Final.Range("G" & i - 13).Value = Final.Range("A" & i - 13).Value
            Final.Range("I" & i - 13).Value = Final.Range("E" & i - 13).Value
            Final.Range("H" & i - 13).Value = "Italy\UniSec\" & Right(Data.Range("G1").Value, 4) & "####"
            Final.Range("M" & i - 13).Value = Data.Range("K" & i).Value
            Final.Range("K" & i - 13).Value = Data.Range("L" & i).Value
            Final.Range("AG" & i - 13).Value = Data.Range("C" & i).Value
            Final.Range("AH" & i - 13).Value = Final.Range("AG" & i - 13).Value
            Final.Range("AI" & i - 13).Value = Final.Range("AG" & i - 13).Value
            Final.Range("AJ" & i - 13).Value = 0
            Final.Range("AK" & i - 13).Value = Data.Range("F" & i).Value
            Final.Range("AL" & i - 13).Value = Final.Range("AK" & i - 13).Value
            Final.Range("AM" & i - 13).Value = Final.Range("AK" & i - 13).Value
            Final.Range("AO" & i - 13).Value = Final.Range("AL" & i - 13).Value
            Final.Range("AN" & i - 13).Value = 1
            Final.Range("AP" & i - 13).Value = 1
            Final.Range("BC" & i - 13).Value = 1
       '----------Cut for ferules-- StrippingLength-----------------
       
      Final.Range("O" & i - 13).Value = Data.Range("T" & i).Value
       Final.Range("BA" & i - 13).Value = 1
        End If
        Next i


        '----------Condition If cell is empty------------------
         Final.Range("A:A").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
         

    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    Sheets("Wiring table").Select
    Range("A15").Select


   lr = Final.Range("A" & Rows.Count).End(xlUp).Row
    
  '---------Изтриване на Sheet1------------------
    Final.Columns("CO").EntireColumn.Delete

    
    Dim wb As Workbook
    Set wb = Workbooks.Add
    ThisWorkbook.Sheets("Komax").Copy Before:=wb.Sheets(1)
    
        '-------------add user in Footer ---------------
    With ActiveSheet.PageSetup
    .LeftFooter = "&D" & Chr(13) & "&9" & Application.UserName
    .RightFooter = "Page " & "&P" & Chr(13) & "&9" & Tools.Label8.Caption
    End With
    
   '---------Изтриване на Sheet1------------------
        
    Application.DisplayAlerts = False
    Sheets("Sheet1").Delete
    Application.DisplayAlerts = True
    
    
Application.CutCopyMode = False 'esp

Dim sFileSaveName As Variant
Dim sPath As String
sPath = ActiveSheet.Range("A2").Value
InitialFoldr$ = "\\10.28.38.5\ppmv\Productions\Italian\LVC\UniSec\!!!__Orders\!_____Ongoing Orders"
sFileSaveName = Application.GetSaveAsFilename(InitialFileName:=Left(sPath, 15), fileFilter:="Excel Files (*.csv), *.xlsm")
If sFileSaveName <> False Then
ActiveWorkbook.SaveAs sFileSaveName, FileFormat:=xlCSV, Local:=True
End If
Else
 answer = MsgBox("To generate Comax table please make worksheet Wiring table active !!!", vbYes + vbQuestion, "")
End If
End Sub

