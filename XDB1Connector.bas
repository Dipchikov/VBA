Attribute VB_Name = "XDB1Connector"
Sub XDB1Connector()


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

                       If cell.Value = "XDB1" And cell(1, 2).Value = 25 And cell(1, 4).Value < motor Then
        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = motor '
        
        End If

        
           If cell.Value = "XDB1" And cell(1, 2).Value = 35 And cell(1, 4).Value < motor Then
        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
       cell(1, 4).Value = motor
        
        End If
                  

          If cell.Value = "XDB1" And cell(1, 2).Value = 40 And cell(1, 4).Value < motor Then
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


                       If cell.Value = "XDB1" And cell(1, 2).Value = 25 And cell(1, 7).Value < motor Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
       cell(1, 7).Value = motor
        
        End If
                
        
        
           
                   If cell.Value = "XDB1" And cell(1, 2).Value = 35 And cell(1, 7).Value < motor Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = motor
        
        End If


          If cell.Value = "XDB1" And cell(1, 2).Value = 40 And cell(1, 7).Value < motor Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = motor
        
        End If

 Next
 
 
      '-------------------------Connections"----------------------------------
      
         XDB1_connectors_number.XDB1_connectors_number
    
    '-------------------------------Clear cells if have crossection------------------------------------------
    
    Set MyPlage = Range("A15:A1000")
        For Each cell In MyPlage
        
         If Not IsEmpty(cell(1, 7).Value) And cell.Value = "XDB1" And Not (cell(1, 4).Value = "XDB1") And Left(cell(1, 4).Value, 3) = "XDB" Then
            cell(1, 7).ClearContents
            cell(1, 8).ClearContents
            cell(1, 9).Value = "Direct connection"
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
            End If
            



        Next
            
End Sub
