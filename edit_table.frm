VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} edit_table 
   Caption         =   "Edit table"
   ClientHeight    =   5508
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   2628
   OleObjectBlob   =   "edit_table.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "edit_table"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton10_Click()
SwapTwoRanges.SwapTwoRanges
End Sub

Private Sub CommandButton13_Click()
Swap_form.Show vbModaless
End Sub

Private Sub CommandButton17_Click()
Swap.Swap
End Sub

Private Sub CommandButton18_Click()
formula.formula
End Sub

Private Sub CommandButton19_Click()
 translate.translate
  
End Sub

Private Sub CommandButton3_Click()
renumber.renumber
End Sub

Private Sub Label8_Click()
    Link = "mailto:hristodipchikov@gmail.com"
    ActiveWorkbook.FollowHyperlink Address:=Link, NewWindow:=True
    Unload Me
End Sub

Private Sub UserForm_Click()

End Sub


Private Sub UserForm_Initialize()
Label8.Caption = Tools.Label8.Caption

    Me.StartUpPosition = 0
    Me.Top = 130
    Me.Left = Application.Left - 50 + Application.Width - Me.Width

End Sub
