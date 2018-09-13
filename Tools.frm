VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Tools 
   Caption         =   "UniSec menu"
   ClientHeight    =   9096
   ClientLeft      =   36
   ClientTop       =   -636
   ClientWidth     =   2532
   OleObjectBlob   =   "Tools.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Tools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub CommandButton10_Click()
Dim Rng1 As Range, rng2 As Range
Dim arr1 As Variant, arr2 As Variant
xTitleId = "Change range"
On Error Resume Next
Set Rng1 = Application.Selection
Set Rng1 = Application.InputBox("Range1:", xTitleId, Rng1.Address, Type:=8)
Rng1.Interior.ColorIndex = 3
 
Set rng2 = Application.InputBox("Range2:", xTitleId, Type:=8)
rng2.Interior.ColorIndex = 15

Application.ScreenUpdating = False
arr1 = Rng1.Value
arr2 = rng2.Value
Rng1.Interior.ColorIndex = 0
rng2.Interior.ColorIndex = 0
Rng1.Value = arr2
rng2.Value = arr1
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With
    
    Legend_of_colours.Legend_of_colours

End Sub

Private Sub CommandButton11_Click()
clear_wiring_table.clear_wiring_table
End Sub

Private Sub CommandButton13_Click()
       UserForm2.Show vbModeless
End Sub

Private Sub CommandButton15_Click()
answer = MsgBox("Do you want to start Legend of colours and Check for Errors before close the form ?", vbYesNo + vbQuestion, "Close the form")
  Select Case answer
    Case vbYes
     Errors.Errors
     Legend_of_colours.Legend_of_colours
     Unload UserForm1
    Case vbNo
      Unload UserForm1
End Select

End Sub

Private Sub CommandButton16_Click()
Error_menu.Show vbModaless
End Sub


Private Sub CommandButton17_Click()
Swap.Swap
Legend_of_colours.Legend_of_colours
    
End Sub

Private Sub CommandButton18_Click()
soft_by_colour.soft_by_colour
End Sub

Private Sub CommandButton19_Click()
Komax_table.Komax_table
End Sub

Private Sub CommandButton20_Click()

edit_table.Show vbModaless
End Sub

Private Sub CommandButton21_Click()
SaveAs.SaveAs
End Sub

Private Sub CommandButton22_Click()
Routing.Routing
sernumerr.sernumerr
End Sub

Private Sub CommandButton23_Click()
FilterCriteria.FilterCriteria
End Sub

Private Sub CommandButton24_Click()
'-------------clear filter------------------------------
 On Error Resume Next
  ActiveSheet.ShowAllData
  ActiveWorkbook.Worksheets("Wiring table").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Wiring table").AutoFilter.Sort.SortFields.Add Key _
        :=Range("A14:A550"), SortOn:=xlSortOnValues, order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Wiring table").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Private Sub CommandButton25_Click()
Number_of_connections.Show vbModaless
End Sub

Private Sub CommandButton4_Click()
 Legend_of_colours.Legend_of_colours
End Sub

Private Sub CommandButton7_Click()
Unload Me
End Sub

Private Sub Frame1_Click()

End Sub

Private Sub Frame2_Click()

End Sub

Private Sub Frame6_Click()

End Sub

Private Sub Label8_Click()

    Link = "http://www.hristodipchikov.tk"
    ActiveWorkbook.FollowHyperlink Address:=Link, NewWindow:=True
    Unload Me

End Sub



Private Sub UserForm_Initialize()

    Label8.Caption = "Copyright Hristo Dipchikov© vR3.3"
    Me.StartUpPosition = 0
    Me.Top = 130
    Me.Left = Application.Left - 50 + Application.Width - Me.Width

End Sub

