Attribute VB_Name = "XdxNumberOfConnections"
Sub XdxNumberOfConnections()

Dim lr As Long
Dim irow As Integer
Dim designation As String
Dim sum As Integer
Dim rng As Range
    Dim rng1 As Range
    lr = Range("A" & Rows.Count).End(xlUp).Row
Set MyPlage = Range("C15:C" & lr)

For Each cell In MyPlage
If Left(cell.Value, 4) = "-XDX" Or Left(cell(1, 4).Value, 4) = "-XDX" Then
If Left(cell.Value, 4) = "-XDX" = True Then
designation = cell.Value
Set rng = cell
End If
If Left(cell(1, 4).Value, 4) = "-XDX" = True Then
designation1 = cell(1, 4).Value
Set rng1 = cell(1, 4)
End If

            sum = 0
            sum1 = 0
            
  On Error Resume Next
   Set MyPlage = Range("C14:C" & lr)
        For Each c In MyPlage
                irow = c.Row
                k = c(1, 4).Value
                If (c.Value = designation Or c(1, 4).Value = designation) And (Cells(irow, 9).Value = "Conductor / wire" Or Cells(irow, 9).Value = "Wire jumper") Then

                    sum = sum + 1
                End If

                If (c.Value = designation1 Or c(1, 4).Value = designation1) And (Cells(irow, 9).Value = "Conductor / wire" Or Cells(irow, 9).Value = "Wire jumper") Then

                    sum1 = sum1 + 1
                End If
            Next c

            If sum > 4 Then
                rng.Interior.ColorIndex = 3
            Else
                rng.Interior.ColorIndex = 0
                End If
             If sum1 > 4 Then
                rng1.Interior.ColorIndex = 3
            Else
                rng1.Interior.ColorIndex = 0
            End If



'MsgBox (designation & "=" & sum)
End If
Next cell


End Sub

