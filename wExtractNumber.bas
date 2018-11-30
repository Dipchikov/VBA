Attribute VB_Name = "wExtractNumber"
Public Function wExtractNumber(sinput) As Double
For i = 1 To Len(sinput)
If IsNumeric(Mid(sinput, i, 1)) Then
result = result & Mid(sinput, i, 1)
End If
Next i
wExtractNumber = result
End Function
