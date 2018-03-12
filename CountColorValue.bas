Attribute VB_Name = "CountColorValue"
Sub CountColorValue()

Dim numbers As Long, lastrow As Long
Dim Rng As Range, c As Range
'Find the last row
lastrow = Range("G" & Rows.Count).End(xlUp).Row
'Set the range you want to search through
Set Rng = Range("G15:I15" & lastrow)

'Iterate through each cell in the range
For Each c In Rng
    'If the interior color is 6 (standard yellow), not blank and not a number
    If c.Font.ColorIndex = 3 Then
        'Add 1 to numbers
        numbers = numbers + 1
    End If
Next c

'Message box with the value of numbers, change to display
'however you'd like
Err = MsgBox("Numbers of fixed errors  is = " & numbers, vbOKOnly + vbCritical)
Range("l10").Value = numbers
End Sub


