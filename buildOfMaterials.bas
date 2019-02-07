Attribute VB_Name = "buildOfMaterials"
Sub buildOfMaterials()

Dim lr As Long
Dim i As Long
Dim irow As Integer
Dim iFinishRow As Integer
Dim lastRowDesignation As Long
Dim Designation As String
Dim jump As Integer
Dim lr_temp As Long
Set finish = Sheets("BOM")
Set Data = Sheets("Wiring table")
Set temp = Sheets("temp")

Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False

'------------------------- REF and BOM----------------------------------

If Error_menu.Ref542.Value = True Then
finish.Range("J17") = "Yes"
Else
finish.Range("J17") = "No"
End If
'------------------------- PXOENIX----------------------------------

If Error_menu.PHOENIX.Value = True Then
finish.Range("J18") = "Yes"
Else
finish.Range("J18") = "No"
End If

finish.Range("E160:E180").ClearContents

 
lr = Data.Range("A" & Rows.Count).End(xlUp).Row


'--------------------Equipment designations-----------------------------------------------
Set Equipment = Union(Data.Range("A15:A" & lr), Data.Range("D15:D" & lr))
lastRowDesignation = finish.Range("L" & Rows.Count).End(xlUp).Row
finish.Range("L2:L" & lastRowDesignation).ClearContents

For Each cell In Equipment
 Designation = cell.Value
 finish.Activate
    lastRowDesignation = finish.Range("L" & Rows.Count).End(xlUp).Row
    Set bomRange = finish.Range("L2:L" & lastRowDesignation)
    With bomRange.Select

         Set sap = Selection.Find(What:=Designation, After:=ActiveCell, LookIn:=xlFormulas, _
         LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
         MatchCase:=False, SearchFormat:=False)

         If sap Is Nothing Then
            iFinishRow = lastRowDesignation + 1
            finish.Range("l" & iFinishRow).Value = Designation
          End If
          
          '---------------------Format ------------------------
          Range("L2:L" & iFinishRow).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
          
          
End With
Next cell





'-------------------------------Jumpers-XDX------------------------------------------
Data.Activate





For irow = 14 To lr + 1
   j = irow

Set cmt = Range("A15:A" & lr)
      On Error Resume Next
For i = 1 To 10
         If (Left(Cells(irow, 1).Value, 3) = "XDX" Or Left(Cells(irow, 1).Value, 4) = "XDI6") And Cells(irow, 9).Value = "Saddle jumper" And Cells(irow, 6).Value = Cells(irow + 1, 3).Value And Cells(irow + 1, 6).Value = Cells(irow + 2, 3).Value And Cells(irow + 2, 6).Value = Cells(irow + 3, 3).Value And Cells(irow + 3, 6).Value = Cells(irow + 4, 3).Value And Cells(irow + 4, 6).Value = Cells(irow + 5, 3).Value And Cells(irow + 5, 6).Value <> Cells(irow + 6, 3).Value Then
        finish.Range("E164").Value = finish.Range("E164").Value + 1
        irow = irow + 6
        End If
         
         
        If (Left(Cells(irow, 1).Value, 3) = "XDX" Or Left(Cells(irow, 1).Value, 4) = "XDI6") And Cells(irow, 9).Value = "Saddle jumper" And Cells(irow, 6).Value = Cells(irow + 1, 3).Value And Cells(irow + 1, 6).Value = Cells(irow + 2, 3).Value And Cells(irow + 2, 6).Value = Cells(irow + 3, 3).Value And Cells(irow + 3, 6).Value = Cells(irow + 4, 3).Value And Cells(irow + 4, 6).Value <> Cells(irow + 5, 3).Value Then
        finish.Range("E164").Value = finish.Range("E164").Value + 1
        irow = irow + 5
        End If
         
        If (Left(Cells(irow, 1).Value, 3) = "XDX" Or Left(Cells(irow, 1).Value, 4) = "XDI6") And Cells(irow, 9).Value = "Saddle jumper" And Cells(irow, 6).Value = Cells(irow + 1, 3).Value And Cells(irow + 1, 6).Value = Cells(irow + 2, 3).Value And Cells(irow + 2, 6).Value = Cells(irow + 3, 3).Value And Cells(irow + 3, 6).Value <> Cells(irow + 4, 3).Value Then
        finish.Range("E163").Value = finish.Range("E163").Value + 1
        irow = irow + 4
        End If
        If (Left(Cells(irow, 1).Value, 3) = "XDX" Or Left(Cells(irow, 1).Value, 4) = "XDI6") And Cells(irow, 9).Value = "Saddle jumper" And Cells(irow, 6).Value = Cells(irow + 1, 3).Value And Cells(irow + 1, 6).Value = Cells(irow + 2, 3).Value And Cells(irow + 2, 6).Value <> Cells(irow + 3, 3).Value Then
        finish.Range("E162").Value = finish.Range("E162").Value + 1
        irow = irow + 3
        End If

        If (Left(Cells(irow, 1).Value, 3) = "XDX" Or Left(Cells(irow, 1).Value, 4) = "XDI6") And Cells(irow, 9).Value = "Saddle jumper" And Cells(irow, 6).Value = Cells(irow + 1, 3).Value And Cells(irow + 1, 6).Value <> Cells(irow + 2, 3).Value Then
         finish.Range("E161").Value = finish.Range("E161").Value + 1
         irow = irow + 2
        
        End If
         If (Left(Cells(irow, 1).Value, 3) = "XDX" Or Left(Cells(irow, 1).Value, 4) = "XDI6") And Cells(irow, 9).Value = "Saddle jumper" And Cells(irow, 6).Value <> Cells(irow + 1, 3).Value Then
         finish.Range("E160").Value = finish.Range("E160").Value + 1
         irow = irow + 1
        End If

Next
'------------------------------------------------------Jumpers-XDI------------------------------------------
        For i = 1 To 10
        If Left(Cells(irow, 1).Value, 3) = "XDI" And Not (Left(Cells(irow, 1).Value, 4) = "XDI6") And Cells(irow, 9).Value = "Saddle jumper" And Cells(irow, 6).Value = Cells(irow + 1, 3).Value And Cells(irow + 1, 6).Value = Cells(irow + 2, 3).Value And Cells(irow + 2, 6).Value = Cells(irow + 3, 3).Value And Cells(irow + 3, 6).Value <> Cells(irow + 4, 3).Value Then
        finish.Range("E168").Value = finish.Range("E168").Value + 1
        irow = irow + 4
        
        End If
        If Left(Cells(irow, 1).Value, 3) = "XDI" And Not (Left(Cells(irow, 1).Value, 4) = "XDI6") And Cells(irow, 9).Value = "Saddle jumper" And Cells(irow, 6).Value = Cells(irow + 1, 3).Value And Cells(irow + 1, 6).Value = Cells(irow + 2, 3).Value And Cells(irow + 2, 6).Value <> Cells(irow + 3, 3).Value Then
        finish.Range("E167").Value = finish.Range("E167").Value + 1
        irow = irow + 3
        End If
        
        
        If Left(Cells(irow, 1).Value, 3) = "XDI" And Not (Left(Cells(irow, 1).Value, 4) = "XDI6") And Cells(irow, 9).Value = "Saddle jumper" And Cells(irow, 6).Value = Cells(irow + 1, 3).Value And Cells(irow + 1, 6).Value <> Cells(irow + 2, 3).Value Then
         finish.Range("E166").Value = finish.Range("E166").Value + 1
         irow = irow + 2
        
        End If
         If Left(Cells(irow, 1).Value, 3) = "XDI" And Not (Left(Cells(irow, 1).Value, 4) = "XDI6") And Cells(irow, 9).Value = "Saddle jumper" And Cells(irow, 6).Value <> Cells(irow + 1, 3).Value Then
         finish.Range("E165").Value = finish.Range("E165").Value + 1
         irow = irow + 1
         
        End If
Next

'------------------------------------------------------Jumpers-XDA-XDV PHOENIX------------------------------------------

If Error_menu.PHOENIX.Value = True Then

For i = 1 To 10
        If (Left(Cells(irow, 1).Value, 3) = "XDA" Or Left(Cells(irow, 1).Value, 3) = "XDV") And Cells(irow, 9).Value = "Saddle jumper" And Cells(irow, 6).Value = Cells(irow + 1, 3).Value And Cells(irow + 1, 6).Value <> Cells(irow + 2, 3).Value Then
         finish.Range("E179").Value = finish.Range("E179").Value + 1
         irow = irow + 2
        End If
         If (Left(Cells(irow, 1).Value, 3) = "XDA" Or Left(Cells(irow, 1).Value, 3) = "XDV") And Cells(irow, 9).Value = "Saddle jumper" And Cells(irow, 6).Value <> Cells(irow + 1, 3).Value Then
         finish.Range("E178").Value = finish.Range("E178").Value + 1
         irow = irow + 1
        End If
Next

End If

'------------------------------------------------------Jumpers-XDA-XDV ABB------------------------------------------

If Error_menu.ABB.Value = True Then

For jump = 1 To 5
'----------------------------------------------Black jumpers repeated twice-- PC8-R1 (2-4-6,7)(9-11-13,14)------------------------
        If (Left(Cells(irow, 1).Value, 3) = "XDA" Or Left(Cells(irow, 1).Value, 3) = "XDV") And Cells(irow, 9).Value = "Insertable jumper" And Cells(irow, 2).Value = 2 And Cells(irow, 5).Value = 4 And Cells(irow + 1, 2).Value = 4 And Cells(irow + 1, 5).Value = 6 And Cells(irow + 2, 2).Value = 6 And Cells(irow + 2, 5).Value = 7 Then
        If Cells(irow + 1, 3).Value = Cells(irow, 6).Value Then
        finish.Range("E174").Value = finish.Range("E174").Value + 1
        irow = irow + 3
        End If
        End If
        
                If (Left(Cells(irow, 1).Value, 3) = "XDA" Or Left(Cells(irow, 1).Value, 3) = "XDV") And Cells(irow, 9).Value = "Insertable jumper" And Cells(irow, 2).Value = 9 And Cells(irow, 5).Value = 11 And Cells(irow + 1, 2).Value = 11 And Cells(irow + 1, 5).Value = 13 And Cells(irow + 2, 2).Value = 13 And Cells(irow + 2, 5).Value = 14 Then
        If Cells(irow + 1, 3).Value = Cells(irow, 6).Value Then
        finish.Range("E174").Value = finish.Range("E174").Value + 1
        irow = irow + 3
        End If
        End If
           
'----------------------------------------------Black jumpers repeated twice-- PC8-R2 (1-4-7,8)(3-6-9,10))------------------------
        
        If (Left(Cells(irow, 1).Value, 3) = "XDA" Or Left(Cells(irow, 1).Value, 3) = "XDV") And Cells(irow, 9).Value = "Insertable jumper" And Cells(irow, 2).Value = 1 And Cells(irow, 5).Value = 4 And Cells(irow + 1, 2).Value = 4 And Cells(irow + 1, 5).Value = 7 And Cells(irow + 2, 2).Value = 6 And Cells(irow + 2, 5).Value = 8 Then
        If Cells(irow + 1, 3).Value = Cells(irow, 6).Value Then
        finish.Range("E175").Value = finish.Range("E175").Value + 1
        irow = irow + 3
        End If
        End If
        
            If (Left(Cells(irow, 1).Value, 3) = "XDA" Or Left(Cells(irow, 1).Value, 3) = "XDV") And Cells(irow, 9).Value = "Insertable jumper" And Cells(irow, 2).Value = 3 And Cells(irow, 5).Value = 6 And Cells(irow + 1, 2).Value = 6 And Cells(irow + 1, 5).Value = 9 And Cells(irow + 2, 2).Value = 9 And Cells(irow + 2, 5).Value = 10 Then
        If Cells(irow + 1, 3).Value = Cells(irow, 6).Value Then
        finish.Range("E175").Value = finish.Range("E175").Value + 1
        irow = irow + 3
        End If
        End If
                    If (Left(Cells(irow, 1).Value, 3) = "XDA" Or Left(Cells(irow, 1).Value, 3) = "XDV") And Cells(irow, 9).Value = "Insertable jumper" And Cells(irow, 2).Value = 13 And Cells(irow, 5).Value = 16 And Cells(irow + 1, 2).Value = 16 And Cells(irow + 1, 5).Value = 19 And Cells(irow + 2, 2).Value = 19 And Cells(irow + 2, 5).Value = 20 Then
        If Cells(irow + 1, 3).Value = Cells(irow, 6).Value Then
        finish.Range("E175").Value = finish.Range("E175").Value + 1
        irow = irow + 3
        End If
        End If
        
        
        
        
           '----------------------------------------------Black jumpers repeated twice-- PC8-R3 (1-4-7-10)(11-14-17-20)------------------------
          If (Left(Cells(irow, 1).Value, 3) = "XDA" Or Left(Cells(irow, 1).Value, 3) = "XDV") And Cells(irow, 9).Value = "Insertable jumper" And Cells(irow, 2).Value = 1 And Cells(irow, 5).Value = 4 And Cells(irow + 1, 2).Value = 4 And Cells(irow + 1, 5).Value = 7 And Cells(irow + 2, 2).Value = 7 And Cells(irow + 2, 5).Value = 10 Then
        If Cells(irow + 1, 3).Value = Cells(irow, 6).Value Then
        finish.Range("E176").Value = finish.Range("E176").Value + 1
        irow = irow + 3
        End If
        End If
                  If (Left(Cells(irow, 1).Value, 3) = "XDA" Or Left(Cells(irow, 1).Value, 3) = "XDV") And Cells(irow, 9).Value = "Insertable jumper" And Cells(irow, 2).Value = 11 And Cells(irow, 5).Value = 14 And Cells(irow + 1, 2).Value = 14 And Cells(irow + 1, 5).Value = 17 And Cells(irow + 2, 2).Value = 17 And Cells(irow + 2, 5).Value = 20 Then
        If Cells(irow + 1, 3).Value = Cells(irow, 6).Value Then
        finish.Range("E176").Value = finish.Range("E176").Value + 1
        irow = irow + 3
        End If
        End If
        
        
        
        
   Next jump
        '---------------------------------------------------------------------------------------------------------------------------------

        If (Left(Cells(irow, 1).Value, 3) = "XDA" Or Left(Cells(irow, 1).Value, 3) = "XDV") And Cells(irow, 9).Value = "Insertable jumper" And Cells(irow, 6).Value = Cells(irow + 1, 3).Value And Cells(irow + 1, 6).Value = Cells(irow + 2, 3).Value Then
         finish.Range("E172").Value = finish.Range("E172").Value + 1
        irow = irow + 3
        End If

        If (Left(Cells(irow, 1).Value, 3) = "XDA" Or Left(Cells(irow, 1).Value, 3) = "XDV") And Cells(irow, 9).Value = "Insertable jumper" And Cells(irow, 6).Value = Cells(irow + 1, 3).Value And Cells(irow + 1, 6).Value <> Cells(irow + 2, 3).Value Then
         finish.Range("E171").Value = finish.Range("E171").Value + 1
         irow = irow + 2
        End If
         If (Left(Cells(irow, 1).Value, 3) = "XDA" Or Left(Cells(irow, 1).Value, 3) = "XDV") And Cells(irow, 9).Value = "Insertable jumper" And Cells(irow, 6).Value <> Cells(irow + 1, 3).Value Then
         finish.Range("E170").Value = finish.Range("E170").Value + 1
         irow = irow + 1
        End If


End If



 k = irow
 If k > j Then
  irow = irow - 1
End If
  
Next irow

Set jumper = finish.Range("E160", "E180")
For Each cell In jumper
If Not IsEmpty(cell.Value) Then

cell.Value = Round(cell.Value * 1.2, 0)
End If

Next
'-----------------------------------------------------------------------Connectors XDA-XDV---------------------------------------------
finish.Range("E130", "E132").Value = 0
finish.Range("E140", "E143").Value = 0

If Error_menu.ABB.Value = True Then

Dim Count_XDA As Double 'Single
Dim Count_XDV As Double 'Single
Dim nameA As String
Dim nameV As String



lr = Range("A" & Rows.Count).End(xlUp).Row
Set MyPlage = Range("D15:d" & lr)

For j = 1 To 10
nameA = "XDA" & j
nameV = "XDV" & j

Count_XDA = WorksheetFunction.CountIf(MyPlage, nameA)
Count_XDV = WorksheetFunction.CountIf(MyPlage, nameV)

If Count_XDA = 2 Then
finish.Range("E130").Value = finish.Range("E130").Value + 1
End If
If Count_XDA > 2 And Count_XDA <= 4 Then
finish.Range("E131").Value = finish.Range("E131").Value + 1
End If
If Count_XDA > 4 And Count_XDA <= 6 Then
finish.Range("E132").Value = finish.Range("E132").Value + 1
End If


'------------------------------------XDV-----------------------

If Count_XDV = 2 And nameV <> "XDV4" Then
finish.Range("E140").Value = finish.Range("E140").Value + 1
End If
If Count_XDV = 2 And nameV = "XDV4" Then
finish.Range("E143").Value = finish.Range("E143").Value + 1
End If
If Count_XDV > 2 And Count_XDA <= 4 Then
finish.Range("E141").Value = finish.Range("E141").Value + 1
End If
If Count_XDV > 4 And Count_XDA <= 6 Then
finish.Range("E142").Value = finish.Range("E142").Value + 1
End If


Next j
End If


'------------------------------------------------------terminlas------------------------------------------

temp.Columns("A:A").ClearContents
Data.Range("C15:C" & lr).Copy
temp.Select

    temp.Range("A1").Select
    'ActiveSheet.Paste
    Range("A1").PasteSpecial Paste:=xlPasteValues
    Data.Select
    
Data.Range("F15:F" & lr).Copy
temp.Select
lr_temp = temp.Range("A" & Rows.Count).End(xlUp).Row
temp.Range("A" & lr_temp + 1).Select

temp.Range("A" & lr_temp + 1).PasteSpecial Paste:=xlPasteValues

temp.Columns("A:A").RemoveDuplicates Columns:=1, Header:=xlNo

Data.Select



Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True

Application.ScreenUpdating = True
End Sub
