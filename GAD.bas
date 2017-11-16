Attribute VB_Name = "GAD"
Sub GAD()

'----------------- minimal wires crossection   --------------------

Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False

Dim wire As String

Set MyPlage = Range("G15:G1000")


wire = InputBox("Please add minimal cross-section of conductors", "Read the General Arrangement Drawings", 1)
For Each cell In MyPlage
 If Not IsEmpty(cell.Value) And cell.Value < wire Then
 cell.Value = wire
 cell.Font.ColorIndex = 3
 cell.Font.Bold = True
End If
Next
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True

End Sub
