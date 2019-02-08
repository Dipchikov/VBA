Attribute VB_Name = "buildOfMaterials"
Sub buildOfMaterials()

Dim lr As Long
Dim i As Long
Dim irow As Integer
Dim iFinishRow As Integer
Dim lastRowDesignation As Long
Dim arrInside As Variant
Dim designation As String
Dim jump As Integer
Dim lr_temp As Long
Dim stoper As Byte
Dim dict As Object
Dim BJMI52p As Byte
Dim BJMI53p As Byte
Dim BJMI54p As Byte
Dim BJMI55p As Byte
Dim BJMI510p As Byte

Dim BJMI82p As Byte
Dim BJMI83p As Byte
Dim BJMI84p As Byte
Dim BJMI85p As Byte
Dim BJMI810p As Byte

Dim XDA2p As Byte
Dim XDA3p As Byte
Dim XDA4p As Byte
Dim XDA10p As Byte

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

Set dict = CreateObject("Scripting.Dictionary")

    For Each CELL In Equipment
      dict(CELL.Value) = CELL
    Next
    

finish.Range("l2").Resize(dict.Count) = Application.Transpose(dict.keys)

finish.Select
          '---------------------Format ------------------------
    lastRowDesignation = finish.Range("L" & Rows.Count).End(xlUp).Row
    finish.Range("L2:L" & lastRowDesignation).Select
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
          
  Dim myArray() As Variant
  myArray = Data.Range("A1:L" & lr).Value
'-------------------------------Stopers--------------------------------------------------------------------
lastRowDesignation = finish.Range("L" & Rows.Count).End(xlUp).Row
arrInside = Array("BT", "KM", "PJ", "PE", "IE", "EA", "BR", "BM", "BX", "TS", "XDB1", "XDT", "XDE", "PFV", "RAD", "FCM", "TB", "XDC", "XDI", "XDX", "XDA", "XDV", "K1", "K2", "K3", "K4", "KA", "KF", "RAA", "TF", "XE", "KLA", "KLT", "QBM", "AA", "RAR")
Set Stopers = finish.Range("L2:L" & lastRowDesignation)
stoper = 4
For i = LBound(arrInside) To UBound(arrInside)


        l = Len(arrInside(i))
        k = arrInside(i)
        For Each CELL In Stopers
        If Left(CELL.Value, l) = k Then
        stoper = stoper + 1
        End If
        
Next CELL
Next i
finish.Range("E186").Value = Round(stoper * 1.1, 0)

'-------------------------------Jumpers-XDX------------------------------------------
Data.Activate







For irow = 1 To lr


Set cmt = Range("A15:A" & lr)

      On Error Resume Next
        Count = 0
         If (Left(myArray(irow, 1), 3) = "XDX" Or Left(myArray(irow, 1), 4) = "XDI6" Or Left(myArray(irow, 1), 4) = "XDI7") And myArray(irow, 9) = "Saddle jumper" Then
            Count = 1
            Do While myArray(irow, 6) = myArray(irow + 1, 3)
            Count = Count + 1
            irow = irow + 1
            Loop
            If Count > 4 Then
            BJMI510p = BJMI510p + 1
            finish.Range("E164").Value = BJMI510p
            End If
            If Count = 4 Then
            BJMI55p = BJMI55p + 1
            finish.Range("E163").Value = BJMI55p
            End If
            If Count = 3 Then
            BJMI54p = BJMI54p + 1
            finish.Range("E162").Value = BJMI54p
            End If
            If Count = 2 Then
            BJMI53p = BJMI53p + 1
            finish.Range("E161").Value = BJMI53p
            End If
            If Count = 1 Then
            BJMI52p = BJMI52p + 1
            finish.Range("E160").Value = BJMI52p
            End If
        End If

'-----------------------XDI Jumpers--------------------------------------------------

         If Left(myArray(irow, 1), 3) = "XDI" And Not (Left(myArray(irow, 1), 4) = "XDI6" Or Left(myArray(irow, 1), 4) = "XDI7") And myArray(irow, 9) = "Saddle jumper" Then
            Count = 1
            Do While myArray(irow, 6) = myArray(irow + 1, 3)
            Count = Count + 1
            irow = irow + 1
            Loop
            If Count > 4 Then
            BJMI810p = BJMI810p + 1
            finish.Range("E169").Value = BJMI810p
            End If
            If Count = 4 Then
            BJMI85p = BJMI85p + 1
            finish.Range("E168").Value = BJMI85p
            End If
            If Count = 3 Then
            BJMI84p = BJMI84p + 1
            finish.Range("E167").Value = BJMI84p
            End If
            If Count = 2 Then
            BJMI83p = BJMI83p + 1
            finish.Range("E166").Value = BJMI83p
            End If
            If Count = 1 Then
            BJMI82p = BJMI82p + 1
            finish.Range("E165").Value = BJMI82p
            End If
        End If
    '------------------------------------------------------Jumpers-XDA-XDV PHOENIX------------------------------------------



 If (Left(myArray(irow, 1), 3) = "XDA" Or Left(myArray(irow, 1), 3) = "XDV") And myArray(irow, 9) = "Saddle jumper" Then
            Count = 1
            Do While myArray(irow, 6) = myArray(irow + 1, 3)
            Count = Count + 1
            irow = irow + 1
            Loop
        If Error_menu.PHOENIX.Value = True Then
            If Count > 2 Then
            XDA10p = XDA10p + 1
            finish.Range("E180").Value = XDA10p
            End If
            If Count = 2 Then
            XDA3p = XDA3p + 1
            finish.Range("E179").Value = XDA3p
            End If
            If Count = 1 Then
            XDA2p = XDA2p + 1
            finish.Range("E178").Value = XDA2p
            End If
        End If
     End If

If irow = lr Then
Exit For
End If
Next irow

'------------------------------------------------------Jumpers-XDA-XDV ABB------------------------------------------
For irow = 14 To lr + 1
j = irow

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

'------------------------------------------------
Set jumper = finish.Range("E160", "E180")
    For Each CELL In jumper
    If Not IsEmpty(CELL.Value) Then
    CELL.Value = Round(CELL.Value * 1.2, 0)
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
Set dict = CreateObject("Scripting.Dictionary")
Set myRange = Union(Data.Range("C15:C" & lr), Data.Range("F15:F" & lr))
    For Each CELL In myRange
      dict(CELL.Value) = CELL
    Next
    
temp.Select
temp.Range("A1").Resize(dict.Count) = Application.Transpose(dict.keys)
Data.Select







End Sub
