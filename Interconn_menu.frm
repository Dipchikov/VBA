VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Interconn_menu 
   Caption         =   "Interconnection menu"
   ClientHeight    =   5112
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   2556
   OleObjectBlob   =   "Interconn_menu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Interconn_menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton11_Click()
Clear_Interconnections.Clear_Interconnections
End Sub

Private Sub CommandButton19_Click()
Comax_table_inter.Comax_table_inter
End Sub

Private Sub CommandButton20_Click()
Unload Me
Interconnections.Show vbModaless
End Sub

Private Sub CommandButton21_Click()
Routing_inter.Routing_inter
End Sub

Private Sub CommandButton22_Click()
SaveAsInter.SaveAsInter
End Sub

Private Sub Label8_Click()

End Sub

Private Sub UserForm_Initialize()
    Label8.Caption = Tools.Label8.Caption
    Me.StartUpPosition = 0
    Me.Top = 130
    Me.Left = Application.Left - 50 + Application.Width - Me.Width

End Sub
Private Sub CommandButton1_Click()

'Close the userform
Unload Me

End Sub






