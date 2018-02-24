Attribute VB_Name = "Swap"
Sub Swap()
Dim FirstAddress As String
    Dim MyArr As Variant
    Dim Rng As Range
    Dim i As Long
    With Application
    'Dim s1 As String, s2 As String
    Dim r1 As Range, r2 As Range
    Dim temp1, temp2
    
    
    On Error Resume Next
        ActiveSheet.ShowAllData
        
        .ScreenUpdating = False
        .EnableEvents = False
    End With

    'Search for a Value Or Values in a range
    'You can also use more values like this Array("ron", "dave")
    MyArr = Array("Swap")

    'Search Column or range
    With ActiveSheet.Columns("K")

        'clear the cells in the column to the right
        '.Offset(0, 1).ClearContents

        For i = LBound(MyArr) To UBound(MyArr)

            'If you want to find a part of the rng.value then use xlPart
            'if you use LookIn:=xlValues it will also work with a
            'formula cell that evaluates to "ron"

            Set Rng = .Find(What:=MyArr(i), _
                            after:=.Cells(.Cells.Count), _
                            LookIn:=xlValues, _
                            Lookat:=xlWhole, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlNext, _
                            MatchCase:=False)
  If Not Rng Is Nothing Then
                FirstAddress = Rng.Address
                Do

                    Set r1 = Rng
                    Set r2 = Rng

    temp1 = r1.Offset(, -10).Resize(, 3).Value
    temp2 = r2.Offset(, -7).Resize(, 3).Value


    r1.Offset(, -10).Resize(, 3).Value = temp2
    r2.Offset(, -7).Resize(, 3).Value = temp1



                    Set Rng = .FindNext(Rng)
                    If Rng Is Nothing Then
                        Exit Do
                    End If
           

                Loop While Rng.Address <> FirstAddress
                
            End If
        Next i
    End With
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With
End Sub
