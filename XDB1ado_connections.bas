Attribute VB_Name = "XDB1ado_connections"
Sub XDB1ado_connections()

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
End Sub
