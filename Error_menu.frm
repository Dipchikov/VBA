VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Error_menu 
   Caption         =   "Common Errors"
   ClientHeight    =   1440
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   6624
   OleObjectBlob   =   "Error_menu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Error_menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CheckBox3_Click()

End Sub

Private Sub CommandButton2_Click()
Application.ScreenUpdating = False
If CheckBox3.Value = True Then
Range("A15:N1000").Interior.ColorIndex = 0
translate.translate
Swap.Swap
Jumpers.Jumpers
Errors.Errors
tfm.tfm
Legend_of_colours.Legend_of_colours
Error_number_of_conections.Error_number_of_conections
'------------------------- Jumpers clear cells"----------------------------------
   



 End If
 
    If CheckBox6.Value = True Then
       FCM3.FCM3
  End If
 
If CheckBox1.Value = True And CheckBox2.Value = True Then

answer = MsgBox("You can select only one option in XDB !!!", vbYesNo + vbQuestion, "Clear the table")
 End If
If CheckBox1.Value = True And CheckBox2.Value = False Then

        XDB1ado.XDB1ado
        
        End If
     If CheckBox2.Value = True And CheckBox1.Value = False Then
       XDB1Connector.XDB1Connector
    
  End If
  
       If CheckBox4.Value = True Then
       ErrorsREf542.ErrorsREf542
    
  End If
         If CheckBox4.Value = False Then
       ErrorsRefs.ErrorsRefs
  End If
CountColorValue.CountColorValue
MsgBox "Now" & vbNewLine & "1. Check Ref numbers of connections" & vbNewLine & "2. Chack all metal jumpes for XDA ,XDV ,XDI,XDX and numbers of conections for them" & vbNewLine & "3. Check all wires sections"
Application.ScreenUpdating = True
Unload Me
End Sub



Private Sub Label8_Click()

End Sub

Private Sub UserForm_Initialize()
Label8.Caption = Tools.Label8.Caption

    Me.StartUpPosition = 0
    Me.Top = 150
    Me.Left = Application.Left - 250 + Application.Width - Me.Width

End Sub
