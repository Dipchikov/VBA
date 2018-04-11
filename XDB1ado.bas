Attribute VB_Name = "XDB1ado"
Sub XDB1ado()

Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
Dim motor As String

'---------------------------motor----------------------------------------------
  motor = InputBox("Please add cross-section of conductors motor circuit" & vbNewLine & "Cross-section of conductors for motor circuit  by default is = 2,5", "Cross-Section for motor circuit", "2,5")

'---------------------------XDB1----------------------------------------------
         Set MyPlage = Range("D15:d1000")
        For Each cell In MyPlage
       
                       If cell.Value = "XDB1" And cell(1, 2).Value = 1 And cell(1, 4).Value < motor Then
        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = motor
        
        End If

                   If cell.Value = "XDB1" And cell(1, 2).Value = 2 And cell(1, 4).Value < motor Then
        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = motor
        End If
        
           
 Next
            
  
     Set MyPlage = Range("A15:A1000")
        For Each cell In MyPlage
        
                      
        If cell.Value = "XDB1" And cell(1, 2).Value = 1 And cell(1, 7).Value < motor Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = motor
        
        End If
                   If cell.Value = "XDB1" And cell(1, 2).Value = 2 And cell(1, 7).Value < motor Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = motor
                  
           End If
           
          
    Next
            
         '---------------------------XDB----------------------------------------------
    
 Set MyPlage = Range("D15:d1000")
  For Each cell In MyPlage
  
        
            If cell.Value = "XDB" And cell(1, 2).Value = 1 And cell(1, 4).Value < motor Then
        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = motor
        
        End If
                   If cell.Value = "XDB" And cell(1, 2).Value = 2 And cell(1, 4).Value < motor Then
        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = motor
                  
         End If
Next



'---------------------------XDB1-XDB is Direct connection---------------------------------------------
         Set MyPlage = Range("A15:A1000")
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
