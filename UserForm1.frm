VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Weidmuller terminals"
   ClientHeight    =   1332
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   6072
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CheckBox2_Click()

End Sub

Private Sub CommandButton2_Click()
Legend_of_feruless.Legend_of_feruless
Unload Me
End Sub

Private Sub UserForm_Click()

End Sub
Private Sub UserForm_Initialize()
Label8.Caption = Tools.Label8.Caption

    Me.StartUpPosition = 0
    Me.Top = 150
    Me.Left = Application.Left - 250 + Application.Width - Me.Width

End Sub
