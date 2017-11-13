Attribute VB_Name = "XDB1ado"
Sub XDB1ado()

Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False


'---------------------------XDB1----------------------------------------------
         Set MyPlage = Range("D15:d1000")
        For Each cell In MyPlage
       
                       If cell.Value = "XDB1" And cell(1, 2).Value = 1 And cell(1, 4).Value < "2,5" Then
        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = "2,5"
        
        End If

                   If cell.Value = "XDB1" And cell(1, 2).Value = 2 And cell(1, 4).Value < "2,5" Then
        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = "2,5"
        End If
        
           
 Next
            
  
     Set MyPlage = Range("A15:A1000")
        For Each cell In MyPlage
        
                      
        If cell.Value = "XDB1" And cell(1, 2).Value = 1 And cell(1, 7).Value < "2,5" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "2,5"
        
        End If
                   If cell.Value = "XDB1" And cell(1, 2).Value = 2 And cell(1, 7).Value < "2,5" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "2,5"
                  
           End If
           
          
    Next
            
         '---------------------------XDB----------------------------------------------
    
 Set MyPlage = Range("D15:d1000")
  For Each cell In MyPlage
  
        
            If cell.Value = "XDB" And cell(1, 2).Value = 1 And cell(1, 4).Value < "2,5" Then
        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = "2,5"
        
        End If
                   If cell.Value = "XDB" And cell(1, 2).Value = 2 And cell(1, 4).Value < "2,5" Then
        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = "2,5"
                  
         End If
Next


      '-------------------------Connections"----------------------------------
   
    Set MyPlage = Range("D15:D1000")
  For Each cell In MyPlage
  
        
        If Left(cell.Value, 4) = "XDB1" And cell(1, 11).Value > 4 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If

Next
   
       Set MyPlage = Range("A15:A1000")
  For Each cell In MyPlage
  
        
        If Left(cell.Value, 4) = "XDB1" And cell(1, 13).Value > 4 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If

 Next
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    
 End Sub
