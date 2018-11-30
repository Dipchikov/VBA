Attribute VB_Name = "XDV_XDA_ABB"
Sub XDV_XDA()

Dim lr As Long
Dim XDA As Double
Dim XDV As Double
lr = Range("A" & Rows.Count).End(xlUp).Row
Set MyPlage = Range("A15:A" & lr)

  For Each cell In MyPlage
        
            If cell.Value = "XDV" And cell.Value = cell(1, 4).Value Then
            answer = MsgBox("Is connection between" & cell(1, 3).Value & " And " & cell(1, 6).Value & " is with-" & cell(1, 9).Value, vbYesNo + vbQuestion, "-XDV jumpers")
            If answer = vbYes Then
            If (cell(1, 9).Value = "Saddle jumper" Or cell(1, 9).Value = "Insertable jumper") Then
            cell(1, 7).ClearContents
            cell(1, 8).ClearContents
            If cell(1, 9).Value = "Insertable jumper" Then
            cell(1, 9).Value = "Saddle jumper"
            cell(1, 9).Font.ColorIndex = 6
            cell(1, 9).Font.Bold = True
            End If
            End If
            If (cell(1, 9).Value = "Wire jumper" Or cell(1, 9).Value = "Conductor / wire") Then
            If Not (cell(1, 9).Value = "Wire jumper") Then
            cell(1, 9).Value = "Wire jumper"
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
            End If
            If IsEmpty(cell(1, 7).Value) Then
            cell(1, 7).Value = XDV
            cell(1, 7).Font.ColorIndex = 3
            cell(1, 7).Font.Bold = True
            End If
            If IsEmpty(cell(1, 8).Value) Then
            cell(1, 8).Value = "bk"
            cell(1, 8).Font.ColorIndex = 3
            cell(1, 8).Font.Bold = True
            End If
            End If
            End If
            
            If answer = vbNo Then
            If (cell(1, 9).Value = "Wire jumper" Or cell(1, 9).Value = "Conductor / wire") Then
            cell(1, 7).ClearContents
            cell(1, 8).ClearContents
            cell(1, 9).Value = "Saddle jumper"
            cell(1, 9).Font.ColorIndex = 6
            cell(1, 9).Font.Bold = True
            Else
             If (cell(1, 9).Value = "Saddle jumper" Or cell(1, 9).Value = "Insertable jumper") Then
            If Not (cell(1, 9).Value = "Wire jumper") Then
            cell(1, 9).Value = "Wire jumper"
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
            End If
            If IsEmpty(cell(1, 7).Value) Then
            cell(1, 7).Value = XDV
            cell(1, 7).Font.ColorIndex = 3
            cell(1, 7).Font.Bold = True
            End If
            If IsEmpty(cell(1, 8).Value) Then
            cell(1, 8).Value = "bk"
            cell(1, 8).Font.ColorIndex = 3
            cell(1, 8).Font.Bold = True
            End If
            End If
            End If

           
            End If
            
 
           

End If

    
        
 Next
 
 
         
'------------------XDA -------------------------
'Range("G7:H1000").Interior.ColorIndex = 0
Set MyPlage = Range("A15:A" & lr)
  For Each cell In MyPlage
  
        
    If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "XDA" And cell(1, 7).Value <> XDA Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = XDA
        End If
        
    Next
    Set MyPlage = Range("D15:D" & lr)
  For Each cell In MyPlage
  
        
            If Not IsEmpty(cell(1, 4).Value) And Left(cell.Value, 3) = "XDA" And cell(1, 4).Value <> XDA Then
        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = XDA
        End If


 Next
         






















End Sub
