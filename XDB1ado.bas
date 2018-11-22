Attribute VB_Name = "XDB1ado"
Sub XDB1ado()

Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False

Dim XDB1 As Double 'Single
XDB1 = motor
Dim lr As Long
'Dim XDB1 As String
lr = Range("A" & Rows.Count).End(xlUp).Row
'---------------------------XDB1----------------------------------------------
'motor = InputBox("Please add cross-section of conductors XDB1 circuit" & vbNewLine & "Cross-section of conductors for XDB1 circuit  by default is = 2,5", "Cross-Section for XDB1 circuit", "2,5")

'---------------------------XDB1----------------------------------------------
        Set MyPlage = Range("D15:d" & lr)
        For Each cell In MyPlage
       
                       If cell.Value = "XDB1" And cell(1, 2).Value = 1 And cell(1, 4).Value <> XDB1 Then
        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = XDB1
        
        End If

                   If cell.Value = "XDB1" And cell(1, 2).Value = 2 And cell(1, 4).Value <> XDB1 Then
        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = XDB1
        End If
        
           
 Next
            
  
     Set MyPlage = Range("A15:A" & lr)
        For Each cell In MyPlage
        
                      
        If cell.Value = "XDB1" And cell(1, 2).Value = 1 And cell(1, 7).Value <> XDB1 Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = XDB1
        
        End If
                   If cell.Value = "XDB1" And cell(1, 2).Value = 2 And cell(1, 7).Value <> XDB1 Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = XDB1
                  
           End If
           
          
    Next
            
         '---------------------------XDB----------------------------------------------
    
 Set MyPlage = Range("D15:d" & lr)
  For Each cell In MyPlage
  
        
            If cell.Value = "XDB" And cell(1, 2).Value = 1 And cell(1, 4).Value <> XDB1 Then
        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = XDB1
        
        End If
                   If cell.Value = "XDB" And cell(1, 2).Value = 2 And cell(1, 4).Value <> XDB1 Then
        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = XDB1
                  
         End If
Next



'---------------------------XDB1-XDB is Direct connection---------------------------------------------
         Set MyPlage = Range("A15:A" & lr)
        For Each cell In MyPlage
       
         If cell.Value = "XDB1" And cell(1, 4).Value = "XDB" And cell(1, 9).Value = "Direct connection" Then
        cell(1, 9).Value = "Conductor / wire"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
       
        
        End If
Next

      '-------------------------Connections"----------------------------------
      
   
        XDB1ado_connections.XDB1ado_connections


    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    
 End Sub
