VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Swap_form 
   Caption         =   "Connectors"
   ClientHeight    =   1080
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   3780
   OleObjectBlob   =   "Swap_form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Swap_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub ComboBox1_Change()

End Sub

Private Sub CommandButton1_Click()
    Dim FirstAddress As String
    Dim MyArr As Variant
    Dim Rng As Range
    Dim i As Long
    With Application
    'Dim s1 As String, s2 As String
    Dim r1 As Range, r2 As Range
    Dim temp1, temp2


        .ScreenUpdating = False
        .EnableEvents = False
    End With

    'Search for a Value Or Values in a range
    'You can also use more values like this Array("ron", "dave")
    MyArr = Array(ComboBox1.Value)

    'Search Column or range
    With ActiveSheet.Columns("A")

        'clear the cells in the column to the right
        '.Offset(0, 1).ClearContents

        For i = LBound(MyArr) To UBound(MyArr)

            'If you want to find a part of the rng.value then use xlPart
            'if you use LookIn:=xlValues it will also work with a
            'formula cell that evaluates to "ron"

            Set Rng = .Find(What:=MyArr(i), _
                            After:=.Cells(.Cells.Count), _
                            LookIn:=xlFormulas, _
                            LookAt:=xlWhole, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlNext, _
                            MatchCase:=False)
  If Not Rng Is Nothing Then
                FirstAddress = Rng.Address
                Do

                    Set r1 = Rng
                    Set r2 = Rng

    temp1 = r1.Offset(, 0).Resize(, 3).Value
    temp2 = r2.Offset(, 3).Resize(, 3).Value


    r1.Offset(, 0).Resize(, 3).Value = temp2
    r2.Offset(, 3).Resize(, 3).Value = temp1



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
    
    Legend_of_colours.Legend_of_colours
    
End Sub




Private Sub Frame1_Click()

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
ComboBox1.List = Array("XDB", "XDB91", "XDB93", "XDB94", "XDB95", "XDB96", "XDB99", "XDH", "XDA1", "XDA2", "XDA3", "XDV1", "XDV2", "XDV3", "XDV4")

End Sub
