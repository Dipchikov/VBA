Attribute VB_Name = "Module3"
Sub xdv()


Dim xdv As Double


'----------default values----------

xdv = XDV1



Dim lr As Long
lr = Range("A" & Rows.Count).End(xlUp).Row

'------------------XDV -terminal------------------------

Set MyPlage = Range("A15:A" & lr)
  For Each cell In MyPlage
  
        
            If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "XDV" And Not (Left(cell(1, 4).Value, 2) = "XE" Or Left(cell(1, 4).Value, 2) = "PE") And cell(1, 7).Value <> xdv Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = xdv
      
        End If
              
        Next
    Set MyPlage = Range("D15:D" & lr)
  For Each cell In MyPlage
    
            If Not IsEmpty(cell(1, 4).Value) And Left(cell.Value, 3) = "XDV" And cell(1, 4).Value <> xdv Then
        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = xdv
        End If
                     
 Next
         
'-------------Black jumpers----------------------------------

Set MyPlage = Range("A15:A" & lr)
  For Each cell In MyPlage
        
        If cell.Value = "XDV" And cell.Value = cell(1, 4).Value And cell(1, 2).Value = 2 And cell(1, 5).Value = 4 Then
        cell(1, 7).ClearContents
        cell(1, 8).ClearContents
        If Not cell(1, 9).Value = "Insertable jumper" Then
        cell(1, 9).Value = "Insertable jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        Else
        If cell.Value = "XDV" And cell.Value = cell(1, 4).Value And cell(1, 2).Value = 4 And cell(1, 5).Value = 6 Then
        cell(1, 7).ClearContents
        cell(1, 8).ClearContents
        If Not cell(1, 9).Value = "Insertable jumper" Then
        cell(1, 9).Value = "Insertable jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If

        Else
        
          If cell.Value = "XDV" And cell.Value = cell(1, 4).Value And cell(1, 2).Value = 4 And cell(1, 5).Value = 6 Then
        cell(1, 7).ClearContents
        cell(1, 8).ClearContents
                If Not cell(1, 9).Value = "Insertable jumper" Then
        cell(1, 9).Value = "Insertable jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If


        Else
                        If cell.Value = "XDV" And cell.Value = cell(1, 4).Value And cell(1, 2).Value = 4 And cell(1, 5).Value = 6 Then
        cell(1, 7).ClearContents
        cell(1, 8).ClearContents
        If Not cell(1, 9).Value = "Insertable jumper" Then
        cell(1, 9).Value = "Insertable jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
       
        Else
        If cell.Value = "XDV" And cell.Value = cell(1, 4).Value And cell(1, 2).Value = 9 And cell(1, 5).Value = 11 Then
        cell(1, 7).ClearContents
        cell(1, 8).ClearContents
        If Not cell(1, 9).Value = "Insertable jumper" Then
        cell(1, 9).Value = "Insertable jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        
        Else
        If cell.Value = "XDV" And cell.Value = cell(1, 4).Value And cell(1, 2).Value = 11 And cell(1, 5).Value = 13 Then
        cell(1, 7).ClearContents
        cell(1, 8).ClearContents
        If Not cell(1, 9).Value = "Insertable jumper" Then
        cell(1, 9).Value = "Insertable jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
       
        Else
         If cell.Value = "XDV" And cell.Value = cell(1, 4).Value And cell(1, 2).Value = 1 And cell(1, 5).Value = 4 Then
        cell(1, 7).ClearContents
        cell(1, 8).ClearContents
        If Not cell(1, 9).Value = "Insertable jumper" Then
        cell(1, 9).Value = "Insertable jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        
        Else
        If cell.Value = "XDV" And cell.Value = cell(1, 4).Value And cell(1, 2).Value = 4 And cell(1, 5).Value = 7 Then
        cell(1, 7).ClearContents
        cell(1, 8).ClearContents
        If Not cell(1, 9).Value = "Insertable jumper" Then
        cell(1, 9).Value = "Insertable jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        
        Else
        
                  If cell.Value = "XDV" And cell.Value = cell(1, 4).Value And cell(1, 2).Value = 3 And cell(1, 5).Value = 6 Then
        cell(1, 7).ClearContents
        cell(1, 8).ClearContents
        If Not cell(1, 9).Value = "Insertable jumper" Then
        cell(1, 9).Value = "Insertable jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        
        Else
               If cell.Value = "XDV" And cell.Value = cell(1, 4).Value And cell(1, 2).Value = 6 And cell(1, 5).Value = 9 Then
        cell(1, 7).ClearContents
        cell(1, 8).ClearContents
        If Not cell(1, 9).Value = "Insertable jumper" Then
        cell(1, 9).Value = "Insertable jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
       
        Else
        
            If cell.Value = "XDV" And cell.Value = cell(1, 4).Value Then
            answer = MsgBox("Is connection between" & cell(1, 3).Value & " And " & cell(1, 6).Value & " is with-" & cell(1, 9).Value, vbYesNo + vbQuestion, "-XDV jumpers")
            If answer = vbYes And (cell(1, 9).Value = "Saddle jumper" Or cell(1, 9).Value = "Insertable jumper") Then
            cell(1, 7).ClearContents
            cell(1, 8).ClearContents
            If cell(1, 9).Value = "Saddle jumper" Then
            cell(1, 9).Value = "Insertable jumper"
            cell(1, 9).Font.ColorIndex = 6
            cell(1, 9).Font.Bold = True
            End If
            End If
            If answer = vbNo Then
            If Not (cell(1, 9).Value = "Wire jumper") Then
            cell(1, 9).Value = "Wire jumper"
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
            End If
            If IsEmpty(cell(1, 7).Value) Then
            cell(1, 7).Value = xdv
            cell(1, 7).Font.ColorIndex = 3
            cell(1, 7).Font.Bold = True
            End If
            If IsEmpty(cell(1, 8).Value) Then
            cell(1, 8).Value = "bk"
            cell(1, 8).Font.ColorIndex = 3
            cell(1, 8).Font.Bold = True
            End If
            
           If answer = vbYes And (cell(1, 9).Value = "Wire jumper" Or cell(1, 9).Value = "Conductor / wire") Then
            If Not (cell(1, 9).Value = "Wire jumper") Then
            cell(1, 9).Value = "Wire jumper"
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
            End If
            If IsEmpty(cell(1, 7).Value) Then
            cell(1, 7).Value = xdv
            cell(1, 7).Font.ColorIndex = 3
            cell(1, 7).Font.Bold = True
            End If
            If IsEmpty(cell(1, 8).Value) Then
            cell(1, 8).Value = "bk"
            cell(1, 8).Font.ColorIndex = 3
            cell(1, 8).Font.Bold = True
            End If
            
            If answer = vbNo And Not (cell(1, 9).Value = "Insertable jumper") Then
            cell(1, 7).ClearContents
            cell(1, 8).ClearContents
            cell(1, 9).Value = "Insertable jumper"
            cell(1, 9).Font.ColorIndex = 6
            cell(1, 9).Font.Bold = True
            End If


End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If

    
        
 Next
 
 


  
        
         


End Sub
