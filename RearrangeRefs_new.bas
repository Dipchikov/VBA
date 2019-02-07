Attribute VB_Name = "RearrangeRefs"
Sub RearrangeRefs()


    Dim ArrRef As Variant
    Dim ArrRefName As Variant
    'Ref615
    Dim ArrConnX5 As Variant
    Dim ArrConn100 As Variant
    Dim ArrConn110 As Variant
    Dim ArrConn120 As Variant
    Dim ArrConn130 As Variant
    '----------Ref620---------------
    Dim ArrConn105 As Variant

    '-------------Ref542--------------
    Dim ArrRefconnectorX10 As Variant
    Dim ArrRefconnectorX20 As Variant
    Dim ArrRefconnectorX21 As Variant
    Dim ArrRefconnectorX30 As Variant
    Dim ArrRefconnectorX31 As Variant
    Dim ArrRefconnectorX40 As Variant
    Dim ArrRefconnectorX41 As Variant
    Dim ArrRefconnectorX50 As Variant
    Dim ArrRefconnectorX60 As Variant
    Dim ArrRefconnectorX80 As Variant
    '-------------Ref601--------------
    Dim ArrConnXK1  As Variant
    Dim ArrConnXK2  As Variant
    Dim ArrConnXK3  As Variant
    Dim ArrConnXK4  As Variant
    Dim ArrConnXK8  As Variant
    Dim ArrConnXK9  As Variant
    Dim ArrConnXK10 As Variant
    
    
    

    Dim rng1 As Range
    Dim rng2 As Range
    Dim temp1 As Range, temp2 As Range
    Dim tempText1, tempText2
    Dim temp As Range
    Dim CellRow As Integer
    Dim NextRow As Long
    Dim name As String
    Dim i As Long
    Dim lr As Long
    Dim j As Long
    Dim l As Long
    Dim k As Variant



    lr = Range("A" & Rows.Count).End(xlUp).Row
    Set myRange = Range("a15:A" & lr)



    On Error Resume Next


    '----------------------------Ref 615 - 620 connectors--------------------------
    '------------------Connector 100-------------------------
    ArrConn100 = Array("100:1", "100:2", "100:3", "100:4", "100:5", "100:6", "100:7", "100:8", "100:9", "100:10", "100:11", "100:12", "100:13", "100:14", "100:15", "100:16", "100:17", "100:18", "100:19", "100:20", "100:21", "100:22", "100:23", "100:24")

    '------------------Connector 100-------------------------
    ArrConn105 = Array("105:1", "105:2", "105:3", "105:4", "105:5", "105:6", "105:7", "105:8", "105:9", "105:10", "105:11", "105:12", "105:13", "105:14", "105:15", "105:16", "105:17", "105:18", "105:19", "105:20", "105:21", "105:22", "105:23", "105:24")

    '------------------Connector 110----------------------------
    ArrConn110 = Array("110:1", "110:2", "110:3", "110:4", "110:5", "110:6", "110:7", "110:8", "110:9", "110:10", "110:11", "110:12", "110:13", "110:14", "110:15", "110:16", "110:17", "110:18", "110:19", "110:20", "110:21", "110:22", "110:23", "110:24")

    '------------------Connector 115----------------------------
    ArrConn115 = Array("115:1", "115:2", "115:3", "115:4", "115:5", "115:6", "115:7", "115:8", "115:9", "115:10", "115:11", "115:12", "115:13", "115:14", "115:15", "115:16", "115:17", "115:18", "115:19", "115:20", "115:21", "115:22", "115:23", "115:24")

    '----------------------Connector 120 ----------------------------
    ArrConn120 = Array("120:1", "120:2", "120:3", "120:4", "120:5", "120:6", "120:7", "120:8", "120:9", "120:10", "120:11", "120:12", "120:13", "120:14")

    '---------------------Connector 130------------------------
    ArrConn130 = Array("130:1", "130:2", "130:3", "130:4", "130:5", "130:6", "130:7", "130:8", "130:9", "130:10", "130:11", "130:12", "130:13", "130:14", "130:15", "130:16", "130:17", "130:18")
 '---------------------Connector X5------------------------
    ArrConnX5 = Array("5:1", "5:2", "5:3", "5:4", "5:5", "5:6", "5:7", "5:8", "5:9", "5:10")

    '----------------------------Ref 542plus - connectors--------------------------
    ArrRefconnectorX10 = Array("X10:1", "X10:2", "X10:3")
    ArrRefconnectorX20 = Array("X20:d2", "X20:z2", "X20:d4", "X20:z4", "X20:d6", "X20:z6", "X20:d8", "X20:z8", "X20:d10", "X20:z10", "X20:d12", "X20:z12", "X20:d14", "X20:z14", "X20:d16", "X20:z16", "X20:d18", "X20:z18", "X20:d20", "X20:z20", "X20:d22", "X20:z22", "X20:d24", "X20:z24", "X20:d26", "X20:z26", "X20:d28", "X20:z28", "X20:d30", "X20:z30")
    ArrRefconnectorX21 = Array("X21:d2", "X21:z2", "X21:d4", "X21:z4", "X21:d6", "X21:z6", "X21:d8", "X21:z8", "X21:d10", "X21:z10", "X21:d12", "X21:z12", "X21:d14", "X21:z14", "X21:d16", "X21:z16", "X21:d18", "X21:z18", "X21:d20", "X21:z20", "X21:d22", "X21:z22", "X21:d24", "X21:z24", "X21:d26", "X21:z26", "X21:d28", "X21:z28", "X21:d30", "X21:z30")
    ArrRefconnectorX30 = Array("X30:d2", "X30:z2", "X30:d4", "X30:z4", "X30:d6", "X30:z6", "X30:d8", "X30:z8", "X30:d10", "X30:z10", "X30:d12", "X30:z12", "X30:d14", "X30:z14", "X30:d16", "X30:z16", "X30:d18", "X30:z18", "X30:d20", "X30:z20", "X30:d22", "X30:z22", "X30:d24", "X30:z24", "X30:d26", "X30:z26", "X30:d28", "X30:z28", "X30:d30", "X30:z30")
    ArrRefconnectorX31 = Array("X31:d2", "X31:z2", "X31:d4", "X31:z4", "X31:d6", "X31:z6", "X31:d8", "X31:z8", "X31:d10", "X31:z10", "X31:d12", "X31:z12", "X31:d14", "X31:z14", "X31:d16", "X31:z16", "X31:d18", "X31:z18", "X31:d20", "X31:z20", "X31:d22", "X31:z22", "X31:d24", "X31:z24", "X31:d26", "X31:z26", "X31:d28", "X31:z28", "X31:d30", "X31:z30")
    ArrRefconnectorX40 = Array("X40:d2", "X40:z2", "X40:d4", "X40:z4", "X40:d6", "X40:z6", "X40:d8", "X40:z8", "X40:d10", "X40:z10", "X40:d12", "X40:z12", "X40:d14", "X40:z14", "X40:d16", "X40:z16", "X40:d18", "X40:z18", "X40:d20", "X40:z20", "X40:d22", "X40:z22", "X40:d24", "X40:z24", "X40:d26", "X40:z26", "X40:d28", "X40:z28", "X40:d30", "X40:z30")
    ArrRefconnectorX41 = Array("X41:d2", "X41:z2", "X41:d4", "X41:z4", "X41:d6", "X41:z6", "X41:d8", "X41:z8", "X41:d10", "X41:z10", "X41:d12", "X41:z12", "X41:d14", "X41:z14", "X41:d16", "X41:z16", "X41:d18", "X41:z18", "X41:d20", "X41:z20", "X41:d22", "X41:z22", "X41:d24", "X41:z24", "X41:d26", "X41:z26", "X41:d28", "X41:z28", "X41:d30", "X41:z30")
    ArrRefconnectorX50 = Array("X50:d2", "X50:z2", "X50:d4", "X50:z4", "X50:d6", "X50:z6", "X50:d8", "X50:z8", "X50:d10", "X50:z10", "X50:d12", "X50:z12", "X50:d14", "X50:z14", "X50:d16", "X50:z16", "X50:d18", "X50:z18", "X50:d20", "X50:z20", "X50:d22", "X50:z22", "X50:d24", "X50:z24", "X50:d26", "X50:z26", "X50:d28", "X50:z28", "X50:d30", "X50:z30")
    ArrRefconnectorX60 = Array("X60:1", "X60:2")
    ArrRefconnectorX80 = Array("X80:1", "X80:2", "X80:3", "X80:4", "X80:5", "X80:6", "X80:7", "X80:8", "X80:9", "X80:10", "X80:11", "X80:12", "X80:13", "X80:14", "X80:15", "X80:16", "X80:17", "X80:18", "X80:19", "X80:20", "X80:21", "X80:22", "X80:23", "X80:24")

'----------------------------Ref 601 - connectors--------------------------
    ArrConnXK1 = Array("XK1:1", "XK1:2", "XK1:3", "XK1:4")
    ArrConnXK2 = Array("XK2:1", "XK2:2", "XK2:3", "XK2:4", "XK2:5", "XK2:6", "XK2:7", "XK2:8", "XK2:9", "XK2:10")
    ArrConnXK3 = Array("XK3:1", "XK3:2", "XK3:3", "XK3:4", "XK3:5")
    ArrConnXK4 = Array("XK4:1", "XK4:2", "XK4:3", "XK4:4")
    ArrConnXK8 = Array("XK8:1", "XK8:2", "XK8:3", "XK8:4", "XK8:5", "XK8:6", "XK8:7", "XK8:8")
    ArrConnXK9 = Array("XK9:1", "XK9:2", "XK9:3", "XK9:4")
    ArrConnXK10 = Array("XK10:1", "XK10:2")
    
            '---------------------Ref protection--arrey------------------------
            ArrRefName = Array("AA", "BCR")
            With myRange
        Application.Calculation = xlCalculationManual
         Application.ScreenUpdating = False
        For s = LBound(ArrRefName) To UBound(ArrRefName)
            name = ArrRefName(i)
            Set rng = .Find(What:=name, _
                            After:=.Cells(.Cells.Count), _
                            LookIn:=xlValues, _
                            LookAt:=xlWhole, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlNext, _
                            MatchCase:=False)
             If Not rng Is Nothing Then
                If name = "AA" Then
                    If Error_menu.Ref542.Value = True Then
                    
                    '----------------------------Ref 542plus ----------------------------
                    ArrRef = Array(ArrRefconnectorX10, ArrRefconnectorX20, ArrRefconnectorX21, ArrRefconnectorX30, ArrRefconnectorX31, ArrRefconnectorX40, ArrRefconnectorX41, ArrRefconnectorX50, ArrRefconnectorX60, ArrRefconnectorX80)
                    Else
                 '----------------------------Ref 615 ----------------------------
                    ArrRef = Array(ArrConn100, ArrConn105, ArrConn110, ArrConn115, ArrConn120, ArrConn130, ArrConnX5)
                    End If
                 End If

                If name = "BCR" Then
                ArrRef = Array(ArrConnXK1, ArrConnXK2, ArrConnXK3, ArrConnXK4, ArrConnXK8, ArrConnXK9, ArrConnXK10)
                name = ArrRefName(i)
                End If
            End If




    NextRow = 14

    For i = LBound(ArrRef) To UBound(ArrRef)
        For j = LBound(ArrRef(i)) To UBound(ArrRef(i))

            l = Len(ArrRef(i)(j))
            k = ArrRef(i)(j)
            For Each cell In myRange
                CellRow = cell.Row
                If cell.Value = name And Right(cell(1, 2).Value, l) = k Then
                    NextRow = NextRow + 1

                     Set rng1 = Range(cell(1, 1), cell(1, 12))
                     Set rng2 = Range(Cells(NextRow, 1), Cells(NextRow, 12))
                    'Set temp = Range(Cells(1, 28), Cells(1, 39))
                    Set temp1 = Range("Y1:AJ1")
                    'rng1.Copy temp
                    
                   'rng2.Copy rng1
                   'temp.Copy rng2
                    'temp.ClearContents
                    'temp.ClearFormats
                    
                    
                    'Swap values
                    rng1.Copy temp1      '.Value '.Offset(, 0)
                    rng2.Copy rng1
                    temp1.Copy rng2
                    temp1.Clear
                    
                    '.Value '.Offset(, 0).
                    'tempText1 = rng1.Font '.Value '.Offset(, 0)
                    'tempText2 = rng2.Font  '.Value '.Offset(, 0).

                   ' rng1 = temp2  '.Offset(, 0).
                   ' rng2 = temp1  '.Offset(, 0).
                    'rng1 = tempText2.Font '.Offset(, 0).
                    'rng2 = tempText1.Font  '.Offset(, 0).

                End If
            Next
        Next j
    Next i
Next s
    
End With

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub


