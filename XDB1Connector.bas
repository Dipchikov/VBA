Attribute VB_Name = "XDB1Connector"
Sub XDB1Connector()



'---------------------------XDB1----------------------------------------------

                   Set MyPlage = Range("D15:d1000")
        For Each cell In MyPlage
       
                       If cell.Value = "XDB1" And cell(1, 2).Value = 1 And cell(1, 4).Value < "2,5" Then
        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = "2,5"
        
        End If

                       If cell.Value = "XDB1" And cell(1, 2).Value = 25 And cell(1, 4).Value < "2,5" Then
        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = "2,5" '
        
        End If

        
           If cell.Value = "XDB1" And cell(1, 2).Value = 35 And cell(1, 4).Value < "2,5" Then
        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
       cell(1, 4).Value = "2,5"
        
        End If
                  

          If cell.Value = "XDB1" And cell(1, 2).Value = 40 And cell(1, 4).Value < "2,5" Then
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


                       If cell.Value = "XDB1" And cell(1, 2).Value = 25 And cell(1, 7).Value < "2,5" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
       cell(1, 7).Value = "2,5"
        
        End If
                
        
        
           
                   If cell.Value = "XDB1" And cell(1, 2).Value = 35 And cell(1, 7).Value < "2,5" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "2,5"
        
        End If


          If cell.Value = "XDB1" And cell(1, 2).Value = 40 And cell(1, 7).Value < "2,5" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "2,5"
        
        End If

 Next
 
 
      '-------------------------Connections"----------------------------------
      
         XDB1_connectors_number.XDB1_connectors_number
    
    '-------------------------------Clear cells if have crossection------------------------------------------
    
    Set MyPlage = Range("A15:A1000")
        For Each cell In MyPlage
        
         If Not IsEmpty(cell(1, 7).Value) And cell.Value = "XDB1" And cell(1, 4).Value = "XDB" Then
            cell(1, 7).ClearContents
            cell(1, 8).ClearContents
            cell(1, 9).Value = "Direct connection"
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
            End If


Next
            
End Sub
