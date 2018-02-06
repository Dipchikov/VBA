VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Routing_form 
   Caption         =   "Routing menu"
   ClientHeight    =   2928
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   2556
   OleObjectBlob   =   "Routing_form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Routing_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton11_Click()
If ActiveSheet.Name = "Routing" Then
 On Error Resume Next
 
    ActiveSheet.ShowAllData
'
' delete Macro
'
answer = MsgBox("Are you sure you want to clear the table? ", vbYesNo + vbQuestion, "Clear the table")
If answer = vbYes Then





Range("A15:L1000").Interior.ColorIndex = 0

   'Range("A4").Select
    'Selection.ClearContents
        Range("B4").Select
    Selection.ClearContents
    Range("A17:e46").Select
    Selection.ClearContents
    Range("A50:E61").Select
    Selection.ClearContents
   
   
    
    Else
    'do nothing
    End If
    

     Range("A17").Select
  End If
End Sub


Private Sub CommandButton21_Click()
SaveAsRouting.SaveAsRouting
End Sub

Private Sub Label8_Click()

End Sub

Private Sub UserForm_Initialize()
    Label8.Caption = Tools.Label8.Caption
    Me.StartUpPosition = 0
    Me.Top = 130
    Me.Left = Application.Left - 50 + Application.Width - Me.Width

End Sub
