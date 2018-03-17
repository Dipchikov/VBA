Attribute VB_Name = "sernumerr"
Sub sernumerr()
' ----дефиниране-------------
Dim shematic As String
Dim err As Single
Dim connections As Single
Dim mydata As Workbook
Dim myserr As Workbook

' ----'присвояване на стойности-------------
Set mydata = Workbooks("Main Italy Secondary table vR2.4.xlsm")
Sheets("Wiring table").Select
err = Range("H10").Value
connections = Range("L10").Value
shematic = Range("B1").Value
 If shematic = "" Then
 rou = MsgBox("Please add scheme number in cell B1!!!", vbExclamation)
  Exit Sub
  End If

If IsFileOpen("\\10.28.38.5\ppmv\Productions\Italian\LVC\UniSec\!!!__Orders\!!!____Serial Numbers__Main Labels\Serial Numbers-Unisec_v6.1.xlsm") Then
Set myserr = Workbooks("Serial Numbers-Unisec_v6.1.xlsm")
myserr.Activate
Sheets("Register").Select
Set MyPlage = Range("E15:E1048576")
    For Each cell In MyPlage
        If cell.Value = shematic Then
        cell(1, 12).Value = connections
        cell(1, 13).Value = err
        End If
        Next
        Else
Set myserr = Workbooks.Open("\\10.28.38.5\ppmv\Productions\Italian\LVC\UniSec\!!!__Orders\!!!____Serial Numbers__Main Labels\Serial Numbers-Unisec_v6.1.xlsm")

myserr.Sheets("Register").Select

Set MyPlage = Range("E15:E1048576")
    For Each cell In MyPlage
        If cell.Value = shematic Then
        cell(1, 12).Value = connections
        cell(1, 13).Value = err
        End If
        Next
        End If
    mydata.Activate
  End Sub
    
