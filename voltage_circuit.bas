Attribute VB_Name = "voltage_circuit"
Public Function XDV1() As Single

'---------------------------XDV----------------------------------------------
XDV1 = InputBox("Please add cross-section for voltage circuit conductors" & vbNewLine & "Cross-section by default for voltage circuit conductors is = 1,5mm2", "Cross-Section for Voltage Circuit", 1.5)
End Function

