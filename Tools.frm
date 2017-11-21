VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Tools 
   Caption         =   "UniSec menu"
   ClientHeight    =   8460
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
If ActiveSheet.Name = "Wiring table" Then
Application.ScreenUpdating = False
Application.DisplayAlerts = False
 On Error Resume Next
ActiveSheet.ShowAllData
'
' delete Macro
'
answer = MsgBox("Are you sure you want to clear the table? Did you press the Routing botton?", vbYesNo + vbQuestion, "Clear the table")
If answer = vbYes Then





Range("A15:N1000").Interior.ColorIndex = 0

    Range("B1").Select
    Selection.ClearContents
        Range("O12").Select
    Selection.ClearContents
    Range("A15:L551").Select
    Selection.ClearContents
    
    Columns("E:E").Select
    Selection.NumberFormat = "@"
    Columns("B:B").Select
    Selection.NumberFormat = "@"
    
   Range("A15:L551").Select
    Range("L551").Activate
    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
    Selection.Font.Bold = False
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
    
    Else
    'do nothing
    End If
    '-------------Formulas---------------
    
    Range("C15").Select
    ActiveCell.FormulaR1C1 = "=""-""&RC[-2]&"":""&RC[-1]"
    Range("C15").Select
    Selection.AutoFill Destination:=Range("C15:C551"), Type:=xlFillDefault
    Range("C15:C551").Select
    Range("F15").Select
    ActiveCell.FormulaR1C1 = "=""-""&RC[-2]&"":""&RC[-1]"
    Range("F15").Select
    Selection.AutoFill Destination:=Range("F15:F551"), Type:=xlFillDefault
    Range("F15:F551").Select
    
    
       '-------------Length formula---------------
        Range("K15").Select
        
    ActiveCell.FormulaR1C1 = _
        "=IF(ISBLANK(RC[-4]),""-"",INDEX(INDIRECT(R12C15),MATCH(RC[-10],'Standard length'!R1C1:R500C1,0),MATCH(RC[-7],'Standard length'!R1C1:R1C500,0)))"
    Range("K15").Select
    Selection.AutoFill Destination:=Range("K15:K551")
    Range("K15:K551").Select
    
     '-------------Cable type formula---------------
     
        Range("L15").Select
    ActiveCell.FormulaR1C1 = _
        "=IFNA(INDEX(INDIRECT(R12C13),MATCH(RC[-4],'Type of cables '!R2C1:R15C1,0),MATCH(RC[-5],'Type of cables '!R2C1:R2C15,0)),""-"")"
    Range("L15").Select
    Selection.AutoFill Destination:=Range("L15:L551")
    Range("L15:L551").Select
    
   '-------------Possible_errors---------------
    Possible_errors.Possible_errors
    

    
     Application.ScreenUpdating = True
Application.DisplayAlerts = True
Range("A15").Select
End If

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
Error_menu.Show vbModeless
End Sub


Private Sub CommandButton17_Click()
Swap.Swap
Legend_of_colours.Legend_of_colours
    
End Sub

Private Sub CommandButton18_Click()
soft_by_colour.soft_by_colour
End Sub

Private Sub CommandButton19_Click()
Comax_table.Comax_table
End Sub

Private Sub CommandButton20_Click()
Unload Me
edit_table.Show vbModaless
End Sub

Private Sub CommandButton21_Click()
SaveAs.SaveAs
End Sub

Private Sub CommandButton22_Click()
Routing.Routing
End Sub

Private Sub CommandButton23_Click()
FilterCriteria.FilterCriteria
End Sub

Private Sub CommandButton24_Click()
  ActiveSheet.ShowAllData
  ActiveWorkbook.Worksheets("Wiring table").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Wiring table").AutoFilter.Sort.SortFields.Add Key _
        :=Range("A14:A550"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Wiring table").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Private Sub CommandButton4_Click()
 Legend_of_colours.Legend_of_colours
End Sub

Private Sub CommandButton7_Click()
Unload Me
End Sub

Private Sub Frame1_Click()

End Sub

Private Sub Frame6_Click()

End Sub

Private Sub Label8_Click()

    Link = "http://www.hristodipchikov.tk"
    ActiveWorkbook.FollowHyperlink Address:=Link, NewWindow:=True
    Unload Me

End Sub



Private Sub UserForm_Initialize()

    Label8.Caption = "created by Hristo Dipchikov © vB8"
    Me.StartUpPosition = 0
    Me.Top = 130
    Me.Left = Application.Left - 50 + Application.Width - Me.Width

End Sub

