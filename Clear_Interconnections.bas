Attribute VB_Name = "Clear_Interconnections"

Sub Clear_Interconnections()

If ActiveSheet.Name = "Interconnections" Then
' delete Macro Interconnections
'

answer = MsgBox("Are you sure you want to clear the table?", vbYesNo + vbQuestion, "Clear the table")
If answer = vbYes Then
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False

Range("A12:J1000").Interior.ColorIndex = 0
     Range("B2").Select
     Selection.ClearContents
     Range("E1").Select
    Selection.ClearContents
     Range("A12:J515").Select
    Selection.ClearContents
'------------------Formulas--------------------------------
    Range("C12:C515").formula = "=""=""&RC[-2]&"":""&RC[-1]"

    Range("F12:F515").formula = "=""=""&RC[-2]&"":""&RC[-1]"

    Range("I12:I515").formula = "=IF(ISBLANK(RC[-8]),""-"",(MID(RC[-5],2,2)-MID(RC[-8],2,2))+1)"

    Range("J12:J515").formula = "=IFNA(INDEX(INDIRECT(R3C12),MATCH(RC[-3],'Type of cables '!R2C1:R20C1,0),MATCH(RC[-2],'Type of cables '!R2C1:R2C20,0)),""-"")"

    
    
    Else
    'do nothing
    End If
    End If
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    Range("A12").Select
    
End Sub

