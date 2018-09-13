Attribute VB_Name = "XDB1_connectors_number"
Sub XDB1_connectors_number()
Dim lr As Long
lr = Range("A" & Rows.Count).End(xlUp).Row
 '-------------------------Connections"----------------------------------
   
  Set MyPlage = Range("D15:D" & lr)
  For Each cell In MyPlage
  
        
        If Left(cell.Value, 4) = "XDB1" And cell(1, 11).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        
        End If

        
        
                If Left(cell.Value, 3) = "XDT" And cell(1, 11).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        
        End If

        
                If Left(cell.Value, 3) = "XDE" And cell(1, 11).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If


           If Left(cell.Value, 4) = "XDB1" And cell(1, 11).Value <= 2 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If
                If Left(cell.Value, 3) = "XDT" And cell(1, 11).Value <= 2 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If
                If Left(cell.Value, 3) = "XDE" And cell(1, 11).Value <= 2 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If
   
   
   Next
   
       Set MyPlage = Range("A15:A" & lr)
  For Each cell In MyPlage
  
        
        If Left(cell.Value, 4) = "XDB1" And cell(1, 13).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If

                If Left(cell.Value, 3) = "XDE" And cell(1, 13).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        

                If Left(cell.Value, 3) = "XDT" And cell(1, 13).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
    
        
        If Left(cell.Value, 4) = "XDB1" And cell(1, 13).Value <= 2 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If
                If Left(cell.Value, 3) = "XDE" And cell(1, 13).Value <= 2 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If
                If Left(cell.Value, 3) = "XDT" And cell(1, 13).Value <= 2 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If
        
    Next
End Sub
