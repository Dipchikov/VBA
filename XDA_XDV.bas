Attribute VB_Name = "XDA_XDV"
Sub XDA_XDV_connections()

Dim lr As Long
lr = Range("A" & Rows.Count).End(xlUp).Row

    '---------------------------------------------XDV------------------------------
Set MyPlage = Range("A15:A" & lr)
  For Each cell In MyPlage
  
If Number_of_connections.PHOENIX.Value = True Then
        If cell.Value = "XDV" And cell(1, 13).Value > 3 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        If cell(1, 4).Value = "XDV" And cell(1, 14).Value > 3 Then
        cell(1, 5).Interior.ColorIndex = 3
        End If
        If cell.Value = "XDV" And cell(1, 13).Value <= 2 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If
        If cell(1, 4).Value = "XDV" And cell(1, 14).Value <= 2 Then
        cell(1, 5).Interior.ColorIndex = 0
        End If
        Else
        If cell(1, 4).Value = "XDV" And cell(1, 14).Value > 4 Then
        cell(1, 5).Interior.ColorIndex = 3
        End If
        
            If cell.Value = "XDV" And cell(1, 13).Value > 4 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
           If cell.Value = "XDV" And cell(1, 13).Value <= 3 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If
        If cell(1, 4).Value = "XDV" And cell(1, 14).Value <= 2 Then
        cell(1, 5).Interior.ColorIndex = 0
        End If
End If

Next
   '---------------------------------------------XDA------------------------------
   
 Set MyPlage = Range("A15:A" & lr)
  For Each cell In MyPlage

If Number_of_connections.PHOENIX.Value = True Then
        If cell.Value = "XDA" And cell(1, 13).Value > 3 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        If cell(1, 4).Value = "XDA" And cell(1, 14).Value > 3 Then
        cell(1, 5).Interior.ColorIndex = 3
        End If
        If cell.Value = "XDA" And cell(1, 13).Value <= 3 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If
        If cell(1, 4).Value = "XDA" And cell(1, 14).Value <= 3 Then
        cell(1, 5).Interior.ColorIndex = 0
        End If
        End If
If Number_of_connections.PHOENIX.Value = False Then
        If cell(1, 4).Value = "XDA" And cell(1, 14).Value > 4 Then
        cell(1, 5).Interior.ColorIndex = 3
        End If
        
            If cell.Value = "XDA" And cell(1, 13).Value > 4 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
           If cell.Value = "XDA" And cell(1, 13).Value <= 4 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If
        If cell(1, 4).Value = "XDA" And cell(1, 14).Value <= 4 Then
        cell(1, 5).Interior.ColorIndex = 0
        End If

 End If
 Next



End Sub
