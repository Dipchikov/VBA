Attribute VB_Name = "motor_circuit"
Public Function motor() As String
'---------------------------XDB1----------------------------------------------
motor = InputBox("Please add cross-section of conductors motor circuit" & vbNewLine & "Cross-section of conductors for motor circuit  by default is = 2,5", "Cross-Section for motor circuit", "2,5")
End Function
