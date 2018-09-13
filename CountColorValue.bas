Attribute VB_Name = "CountColorValue"
Sub CountColorValue()

Dim numbers As Long, lastrow As Long
Dim mark As Long
Dim mark2 As Long
Dim mark3 As Long
Dim rng As Range, c As Range
'Find the last row
lastrow = Range("A" & Rows.Count).End(xlUp).Row
'Set the range you want to search through
Set rng = Range("G15:I15" & lastrow)

'Iterate through each cell in the range
For Each c In rng
    'If the interior color is 6 (standard yellow), not blank and not a number
    If c.Font.ColorIndex = 3 Then
        'Add 1 to numbers
        numbers = numbers + 1
    End If
Next c
Set rng = Range("D15:F" & lastrow)

For Each c In rng
    'If the interior color is 6 (standard yellow), not blank and not a number
    If c.Font.ColorIndex = 3 Then
        'Add 1 to numbers
        mark = mark + 1
    End If
Next c

Set rng = Range("A15:C" & lastrow)

For Each c In rng
    'If the interior color is 6 (standard yellow), not blank and not a number
    If c.Font.ColorIndex = 3 Then
        'Add 1 to numbers
        mark2 = mark2 + 1
    End If
Next c
Set rng = Range("I15:I" & lastrow)

For Each c In rng
    'If the interior color is 6 (standard yellow), not blank and not a number
    If c.Font.ColorIndex = 6 Then
        'Add 1 to numbers
        mark3 = mark3 + 1
    End If
Next c

'Message box with the value of numbers, change to display
'however you'd like
err = MsgBox("Numbers of fixed errors  is = " & numbers + mark + mark2 + mark3, vbOKOnly + vbCritical)
Range("H10").Value = numbers + mark + mark2

End Sub


