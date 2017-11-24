VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Number_of_connections 
   Caption         =   "Number of connections"
   ClientHeight    =   1440
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   5868
   OleObjectBlob   =   "Number_of_connections.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Number_of_connections"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton2_Click()
Application.ScreenUpdating = False
If CheckBox3.Value = True Then
Range("A15:F1000").Interior.ColorIndex = 0
Error_number_of_conections.Error_number_of_conections
'------------------------- Jumpers clear cells"----------------------------------
End If
 
If CheckBox1.Value = True And CheckBox2.Value = True Then

answer = MsgBox("You can select only one option in XDB !!!", vbYesNo + vbQuestion, "Clear the table")
End If
 
If CheckBox1.Value = True And CheckBox2.Value = False Then

        XDB1ado_connections.XDB1ado_connections
        
        End If
     If CheckBox2.Value = True And CheckBox1.Value = False Then
       XDB1_connectors_number.XDB1_connectors_number
    
  End If
  
       If CheckBox4.Value = True Then
       ErrorsREf542.ErrorsREf542
    
  End If
         If CheckBox4.Value = False Then
       ErrorsRefs.ErrorsRefs
       
  End If
Application.ScreenUpdating = True
Unload Me
End Sub


Private Sub UserForm_Initialize()
Label8.Caption = Tools.Label8.Caption
    Me.StartUpPosition = 0
    Me.Top = 130
    Me.Left = Application.Left - 50 + Application.Width - Me.Width

End Sub
