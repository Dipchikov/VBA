VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Interconnections 
   Caption         =   "Interconnection Form"
   ClientHeight    =   8148
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   7908
   OleObjectBlob   =   "Interconnections.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Interconnections"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CheckBox2_Click()

End Sub

Private Sub ComboBox1_Change()

End Sub

Private Sub ComboBox2_Change()

End Sub

Private Sub ComboBox8_Change()

End Sub

Private Sub CommandButton1_Click()

Dim iRow As Long
Dim ws As Worksheet
Set ws = ActiveWorkbook.ActiveSheet

'find first empty row in database
'iRow = ws.Cells.Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).Row + 1
iRow = ws.Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row
'check for a Name number
If Trim(Me.TextBox1.Value) = "" Then
Me.TextBox1.SetFocus
MsgBox "Please complete the form"
Exit Sub
End If

If CheckBox2.Value = True And CheckBox7.Value = True Then
'copy the data to the database
ws.Cells(iRow, 1).Value = Me.TextBox1.Value & ":" & Me.ComboBox5.Value
ws.Cells(iRow, 2).Value = Me.CheckBox2.Caption
'ws.Cells(iRow, 3).Value = " = " & Me.TextBox1.Value & ":" & Me.ComboBox5.Value & ":" & Me.CheckBox1.Caption
ws.Cells(iRow, 4).Value = Me.TextBox4.Value & ":" & Me.ComboBox4.Value
ws.Cells(iRow, 5).Value = Me.CheckBox7.Caption
'ws.Cells(iRow, 6).Value = " = " & Me.TextBox4.Value & ":" & Me.ComboBox4.Value & ":" & Me.CheckBox2.Caption
ws.Cells(iRow, 7).Value = Me.ComboBox1.Value
ws.Cells(iRow, 8).Value = Me.ComboBox2.Value
iRow = iRow + 1
End If

If CheckBox4.Value = True And CheckBox9.Value = True Then

ws.Cells(iRow, 1).Value = Me.TextBox1.Value & ":" & Me.ComboBox5.Value
ws.Cells(iRow, 2).Value = Me.CheckBox4.Caption
'ws.Cells(iRow, 3).Value = " = " & Me.TextBox1.Value & ":" & Me.ComboBox5.Value & ":" & Me.CheckBox4.Caption
ws.Cells(iRow, 4).Value = Me.TextBox4.Value & ":" & Me.ComboBox4.Value
ws.Cells(iRow, 5).Value = Me.CheckBox9.Caption
'ws.Cells(iRow, 6).Value = " = " & Me.TextBox4.Value & ":" & Me.ComboBox4.Value & ":" & Me.CheckBox9.Caption
ws.Cells(iRow, 7).Value = Me.ComboBox1.Value
ws.Cells(iRow, 8).Value = Me.ComboBox2.Value
iRow = iRow + 1

End If
If CheckBox6.Value = True And CheckBox11.Value = True Then
ws.Cells(iRow, 1).Value = Me.TextBox1.Value & ":" & Me.ComboBox5.Value
ws.Cells(iRow, 2).Value = Me.CheckBox6.Caption
'ws.Cells(iRow, 3).Value = " = " & Me.TextBox1.Value & ":" & Me.ComboBox5.Value & ":" & Me.CheckBox6.Caption
ws.Cells(iRow, 4).Value = Me.TextBox4.Value & ":" & Me.ComboBox4.Value
ws.Cells(iRow, 5).Value = Me.CheckBox11.Caption
'ws.Cells(iRow, 6).Value = " = " & Me.TextBox4.Value & ":" & Me.ComboBox4.Value & ":" & Me.CheckBox10.Caption
ws.Cells(iRow, 7).Value = Me.ComboBox1.Value
ws.Cells(iRow, 8).Value = Me.ComboBox2.Value
iRow = iRow + 1
End If


'Second Terminal

If CheckBox14.Value = True And CheckBox19.Value = True Then
'copy the data to the database
ws.Cells(iRow, 1).Value = Me.TextBox1.Value & ":" & Me.ComboBox3.Value
ws.Cells(iRow, 2).Value = Me.CheckBox14.Caption
'ws.Cells(iRow, 3).Value = " = " & Me.TextBox1.Value & ":" & Me.ComboBox3.Value & ":" & Me.CheckBox14.Caption
ws.Cells(iRow, 4).Value = Me.TextBox4.Value & ":" & Me.ComboBox6.Value
ws.Cells(iRow, 5).Value = Me.CheckBox19.Caption
'ws.Cells(iRow, 6).Value = " = " & Me.TextBox4.Value & ":" & Me.ComboBox6.Value & ":" & Me.CheckBox19.Caption
ws.Cells(iRow, 7).Value = Me.ComboBox7.Value
ws.Cells(iRow, 8).Value = Me.ComboBox8.Value
iRow = iRow + 1

End If

If CheckBox16.Value = True And CheckBox21.Value = True Then

ws.Cells(iRow, 1).Value = Me.TextBox1.Value & ":" & Me.ComboBox3.Value
ws.Cells(iRow, 2).Value = Me.CheckBox16.Caption
'ws.Cells(iRow, 3).Value = " = " & Me.TextBox1.Value & ":" & Me.ComboBox3.Value & ":" & Me.CheckBox16.Caption
ws.Cells(iRow, 4).Value = Me.TextBox4.Value & ":" & Me.ComboBox6.Value
ws.Cells(iRow, 5).Value = Me.CheckBox21.Caption
'ws.Cells(iRow, 6).Value = " = " & Me.TextBox4.Value & ":" & Me.ComboBox6.Value & ":" & Me.CheckBox21.Caption
ws.Cells(iRow, 7).Value = Me.ComboBox7.Value
ws.Cells(iRow, 8).Value = Me.ComboBox8.Value
iRow = iRow + 1
'MsgBox "Data added", vbOKOnly + vbInformation, "Data Added"
'Me.TextBox1.SetFocus
End If

If CheckBox18.Value = True And CheckBox23.Value = True Then
ws.Cells(iRow, 1).Value = Me.TextBox1.Value & ":" & Me.ComboBox3.Value
ws.Cells(iRow, 2).Value = Me.CheckBox18.Caption
'ws.Cells(iRow, 3).Value = " = " & Me.TextBox1.Value & ":" & Me.ComboBox3.Value & ":" & Me.CheckBox18.Caption
ws.Cells(iRow, 4).Value = Me.TextBox4.Value & ":" & Me.ComboBox6.Value
ws.Cells(iRow, 5).Value = Me.CheckBox23.Caption
'ws.Cells(iRow, 6).Value = " = " & Me.TextBox4.Value & ":" & Me.ComboBox6.Value & ":" & Me.CheckBox23.Caption
ws.Cells(iRow, 7).Value = Me.ComboBox7.Value
ws.Cells(iRow, 8).Value = Me.ComboBox8.Value

iRow = iRow + 1
End If
If TextBox5 = Empty And TextBox6 = Empty Then
 Else
ws.Cells(iRow, 1).Value = Me.TextBox1.Value & ":" & Me.ComboBox11.Value
ws.Cells(iRow, 2).Value = TextBox5.Value
'ws.Cells(iRow, 3).Value = " = " & Me.TextBox1.Value & ":" & Me.ComboBox3.Value & ":" & Me.CheckBox18.Caption
ws.Cells(iRow, 4).Value = Me.TextBox4.Value & ":" & Me.ComboBox12.Value
ws.Cells(iRow, 5).Value = TextBox6.Value
'ws.Cells(iRow, 6).Value = " = " & Me.TextBox4.Value & ":" & Me.ComboBox6.Value & ":" & Me.CheckBox23.Caption
ws.Cells(iRow, 7).Value = Me.ComboBox9.Value
ws.Cells(iRow, 8).Value = Me.ComboBox10.Value
End If
MsgBox "Data added", vbOKOnly + vbInformation, "Data Added"
Me.TextBox1.SetFocus
End Sub



Private Sub CommandButton2_Click()
'clear the data
Me.TextBox1.Value = ""

Me.TextBox4.Value = ""
Me.TextBox5.Value = ""
Me.TextBox6.Value = ""
Me.ComboBox1.Value = ""

Me.ComboBox2.Value = ""
ComboBox3.Value = ""
ComboBox4.Value = ""
ComboBox5.Value = ""
ComboBox6.Value = ""
ComboBox7.Value = ""
ComboBox8.Value = ""
End Sub

Private Sub CommandButton3_Click()
Unload Me
Interconn_menu.Show vbModeless
End Sub



Private Sub Frame1_Click()

End Sub

Private Sub Label6_Click()

End Sub

Private Sub Frame3_Click()

End Sub

Private Sub Label10_Click()

End Sub

Private Sub Label4_Click()

End Sub

Private Sub Label7_Click()

End Sub


Private Sub UserForm_Initialize()
ComboBox1.List = Array("bk", "gnye", "bn", "gr", "bu")
ComboBox7.List = ComboBox1.List
ComboBox9.List = ComboBox1.List
ComboBox2.List = Array("0,2", "0,5", "0,8", "1", "1,5", "2,5", "4", "6")
ComboBox8.List = ComboBox2.List
ComboBox10.List = ComboBox2.List
ComboBox3.List = Array("XDI", "XDI1", "XDI2", "XDI3", "XDI4", "XDI5", "XDI6", "XDI7", "XDI8", "XDI9")
ComboBox4.List = ComboBox3.List
ComboBox5.List = ComboBox3.List
ComboBox11.List = ComboBox3.List
ComboBox6.List = ComboBox3.List
ComboBox12.List = ComboBox3.List
CheckBox1.Value = False
CheckBox2.Value = False
CheckBox3.Value = False
CheckBox4.Value = False
CheckBox5.Value = False
CheckBox6.Value = False
CheckBox7.Value = False
CheckBox8.Value = False
CheckBox9.Value = False
CheckBox10.Value = False
CheckBox11.Value = False
CheckBox12.Value = False
CheckBox13.Value = False
CheckBox14.Value = False
CheckBox15.Value = False
CheckBox16.Value = False
CheckBox17.Value = False
CheckBox18.Value = False
CheckBox19.Value = False
CheckBox20.Value = False
CheckBox21.Value = False
CheckBox22.Value = False
CheckBox23.Value = False
CheckBox24.Value = False
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    If CloseMode = 0 Then
        Cancel = True
        MsgBox "The X is disabled, please use the Close Command Button.", vbCritical
    End If
    
End Sub
