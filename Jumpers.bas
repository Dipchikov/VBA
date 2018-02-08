Attribute VB_Name = "Jumpers"
Sub Jumpers()

Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False

 '------------------Jumpers Primery -------------------------
 
Set MyPlage = Range("G15:G1000")

    For Each cell In MyPlage

   If cell.Value = "Bridge" Then
            cell(1, 3).Value = "Insertable jumper"
            cell.Value = ""
    End If
    Next

'----------------- minimal wires crossection   --------------------
Dim wire As String

Set MyPlage = Range("G15:G1000")


wire = InputBox("Please add minimal cross-section of conductors", "Read the General Arrangement Drawings!!!", 1)
For Each cell In MyPlage
 If Not IsEmpty(cell.Value) And cell.Value < wire Then
 If Not (cell(1, 6).Value = "-" Or cell(1, 6).Value = "Shielded cable") Then
 cell.Value = wire
 cell.Font.ColorIndex = 3
 cell.Font.Bold = True
End If
End If

Next




'---------------------------clear cells BAT-FCF -QAB -BGT -QCE -BCT- BCN- BAD- EB- EA- BGB-QBM2 - BPS--------------------------------

Set MyPlage = Range("A15:A1000")

    For Each cell In MyPlage

        If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "BAT" Then
            cell(1, 7).ClearContents
            cell(1, 8).ClearContents
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
          End If
          
                  If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "BGB" Then
            cell(1, 7).ClearContents
            cell(1, 8).ClearContents
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
          End If

                If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "QCE" Then
            cell(1, 7).ClearContents
            cell(1, 8).ClearContents
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
          End If
          
        If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "FCF" Then
            cell(1, 7).ClearContents
            cell(1, 8).ClearContents
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True

        End If
        
        If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "QAB" Then
            cell(1, 7).ClearContents
            cell(1, 8).ClearContents
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
        End If
    
        If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "BGT" Then
            cell(1, 7).ClearContents
            cell(1, 8).ClearContents
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
        End If
        
                If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "BGE" Then
            cell(1, 7).ClearContents
            cell(1, 8).ClearContents
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
        End If
        
                If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "BCT" Then
            cell(1, 7).ClearContents
            cell(1, 8).ClearContents
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
        End If
               If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "BCN" Then
            cell(1, 7).ClearContents
            cell(1, 8).ClearContents
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
        End If
                 If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "BAD" Then
            cell(1, 7).ClearContents
            cell(1, 8).ClearContents
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
        End If
        
         If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 2) = "EB" Then
            cell(1, 7).ClearContents
            cell(1, 8).ClearContents
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
        End If
        
         If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "EA2" Then
         cell(1, 7).ClearContents
        cell(1, 8).ClearContents
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
           
        If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "BPS" Then
         cell(1, 7).ClearContents
        cell(1, 8).ClearContents
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
       Next
                 
                 
 Set MyPlage = Range("D15:D1000")

    For Each cell In MyPlage
    

          If Not IsEmpty(cell(1, 4).Value) And Left(cell.Value, 3) = "BAT" Then
          
            cell(1, 4).ClearContents
            cell(1, 5).ClearContents
            cell(1, 6).Font.ColorIndex = 3
            cell(1, 6).Font.Bold = True

        End If
        
        If Not IsEmpty(cell(1, 4).Value) And Left(cell.Value, 3) = "FCF" Then
            cell(1, 4).ClearContents
            cell(1, 5).ClearContents
            cell(1, 6).Font.ColorIndex = 3
            cell(1, 6).Font.Bold = True
        End If
            If Not IsEmpty(cell(1, 4).Value) And Left(cell.Value, 3) = "QAB" Then
            cell(1, 4).ClearContents
            cell(1, 5).ClearContents
            cell(1, 6).Font.ColorIndex = 3
            cell(1, 6).Font.Bold = True
        End If

            If Not IsEmpty(cell(1, 4).Value) And Left(cell.Value, 3) = "BGT" Then
            cell(1, 4).ClearContents
            cell(1, 5).ClearContents
            cell(1, 6).Font.ColorIndex = 3
            cell(1, 6).Font.Bold = True
        End If
            If Not IsEmpty(cell(1, 4).Value) And Left(cell.Value, 3) = "BGE" Then
            cell(1, 4).ClearContents
            cell(1, 5).ClearContents
            cell(1, 6).Font.ColorIndex = 3
            cell(1, 6).Font.Bold = True
        End If
                    If Not IsEmpty(cell(1, 4).Value) And Left(cell.Value, 3) = "QCE" Then
            cell(1, 4).ClearContents
            cell(1, 5).ClearContents
            cell(1, 6).Font.ColorIndex = 3
            cell(1, 6).Font.Bold = True
        End If
        
        Next
 '------------------Jumpers between equipment -------------------------
 
Set MyPlage = Range("A15:A1000")

    For Each cell In MyPlage

   If cell.Value <> cell(1, 4).Value And cell(1, 9).Value = "Ponticello a staffa" Then
    x = InputBox("Please add cross-section of conductors between" & vbNewLine & cell(1, 3) & " and " & cell(1, 6), "Wire jumper  between " & cell(1, 3) & " and " & cell(1, 6), wire)
           cell(1, 9).Value = "Conduttore/filo"
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
            If IsEmpty(cell(1, 7).Value) Then
            cell(1, 7).Value = x
            cell(1, 7).Font.ColorIndex = 3
            cell(1, 7).Font.Bold = True
    End If
     If IsEmpty(cell(1, 8).Value) Then
    cell(1, 8).Value = "bk"
    cell(1, 8).Font.ColorIndex = 3
    cell(1, 8).Font.Bold = True
    End If
    End If

    
       If cell.Value <> cell(1, 4).Value And cell(1, 9).Value = "Ponticello inseribile" Then
        x = InputBox("Please add cross-section of conductors between" & vbNewLine & cell(1, 3) & " and " & cell(1, 6), "Wire jumper  between " & cell(1, 3) & " and " & cell(1, 6), wire)
            cell(1, 9).Value = "Conduttore/filo"
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
            If IsEmpty(cell(1, 7).Value) Then
            cell(1, 7).Value = x
    End If
        If IsEmpty(cell(1, 8).Value) Then
    cell(1, 8).Value = "bk"
    cell(1, 8).Font.ColorIndex = 3
    cell(1, 8).Font.Bold = True
    End If
    End If
    
        If cell.Value <> cell(1, 4).Value And cell(1, 9).Value = "Ponticello a filo" Then
         x = InputBox("Please add cross-section of conductors between" & vbNewLine & cell(1, 3) & " and " & cell(1, 6), "Wire jumper  between " & cell(1, 3) & " and " & cell(1, 6), wire)
            cell(1, 9).Value = "Conduttore/filo"
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
            If IsEmpty(cell(1, 7).Value) Then
            cell(1, 7).Value = x
            cell(1, 7).Font.ColorIndex = 3
            cell(1, 7).Font.Bold = True
    End If
        If IsEmpty(cell(1, 8).Value) Then
    cell(1, 8).Value = "bk"
    cell(1, 8).Font.ColorIndex = 3
    cell(1, 8).Font.Bold = True
    End If
    End If
        
            If cell.Value <> cell(1, 4).Value And cell(1, 9).Value = "Insertable jumper" Then
             x = InputBox("Please add cross-section of conductors between" & vbNewLine & cell(1, 3) & " and " & cell(1, 6), "Wire jumper  between " & cell(1, 3) & " and " & cell(1, 6), wire)
            cell(1, 9).Value = "Conductor / wire"
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
             If IsEmpty(cell(1, 7).Value) Then
            cell(1, 7).Value = x
            cell(1, 7).Font.ColorIndex = 3
            cell(1, 7).Font.Bold = True
    End If
        If IsEmpty(cell(1, 8).Value) Then
    cell(1, 8).Value = "bk"
    cell(1, 8).Font.ColorIndex = 3
    cell(1, 8).Font.Bold = True
    End If
    End If
      
                 If cell.Value <> cell(1, 4).Value And cell(1, 9).Value = "Saddle jumper" Then
            x = InputBox("Please add cross-section of conductors between" & vbNewLine & cell(1, 3) & " and " & cell(1, 6), "Wire jumper  between " & cell(1, 3) & " and " & cell(1, 6), wire)
            cell(1, 9).Value = "Conductor / wire"
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
            If IsEmpty(cell(1, 7).Value) Then
            cell(1, 7).Value = x
            cell(1, 7).Font.ColorIndex = 3
            cell(1, 7).Font.Bold = True
    End If
    If IsEmpty(cell(1, 8).Value) Then
    cell(1, 8).Value = "bk"
    cell(1, 8).Font.ColorIndex = 3
    cell(1, 8).Font.Bold = True
    End If
    End If
         
                If cell.Value <> cell(1, 4).Value And cell(1, 9).Value = "Wire jumper" Then
            x = InputBox("Please add cross-section of conductors between" & vbNewLine & cell(1, 3) & " and " & cell(1, 6), "Wire jumper  between " & cell(1, 3) & " and " & cell(1, 6), wire)
            cell(1, 9).Value = "Conductor / wire"
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
            If IsEmpty(cell(1, 7).Value) Then
            cell(1, 7).Value = x
            cell(1, 7).Font.ColorIndex = 3
            cell(1, 7).Font.Bold = True
            End If
    If IsEmpty(cell(1, 8).Value) Then
             cell(1, 8).Value = "bk"
             cell(1, 8).Font.ColorIndex = 3
             cell(1, 8).Font.Bold = True
    End If
    End If
                If cell.Value <> cell(1, 4).Value And cell(1, 9).Value = "Ponticello a filo" Then
            x = InputBox("Please add cross-section of conductors between" & vbNewLine & cell(1, 3) & " and " & cell(1, 6), "Wire jumper  between " & cell(1, 3) & " and " & cell(1, 6), wire)
            cell(1, 9).Value = "Conduttore/filo"
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
            If IsEmpty(cell(1, 7).Value) Then
            cell(1, 7).Value = x
            cell(1, 7).Font.ColorIndex = 3
            cell(1, 7).Font.Bold = True
            End If
    If IsEmpty(cell(1, 8).Value) Then
             cell(1, 8).Value = "bk"
             cell(1, 8).Font.ColorIndex = 3
             cell(1, 8).Font.Bold = True
    End If
    End If
    
 Next
        
        
 '------------------XDA -------------------------

Set MyPlage = Range("A15:A1000")

    For Each cell In MyPlage

        
           If Left(cell.Value, 3) = "XDA" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) >= 1 And cell(1, 9).Value = "Ponticello a staffa" Then
        cell(1, 7).ClearContents
        cell(1, 8).ClearContents
        End If
        End If
   If Left(cell.Value, 3) = "XDA" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) >= 1 And cell(1, 9).Value = "Ponticello inseribile" Then
         cell(1, 7).ClearContents
            cell(1, 8).ClearContents

        End If
        End If
        
           If Left(cell.Value, 3) = "XDA" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) >= 1 And cell(1, 9).Value = "Insertable jumper" Then
        cell(1, 7).ClearContents
        cell(1, 8).ClearContents
        End If
        End If
   If Left(cell.Value, 3) = "XDA" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) >= 1 And cell(1, 9).Value = "Saddle jumper" Then
        cell(1, 7).ClearContents
        cell(1, 8).ClearContents
        End If
        End If
        
 Next
        
'------------------XDV -------------------------

Set MyPlage = Range("A15:A1000")

    For Each cell In MyPlage

        
           If Left(cell.Value, 3) = "XDV" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) >= 1 And cell(1, 9).Value = "Ponticello a staffa" Then
        cell(1, 7).ClearContents
        cell(1, 8).ClearContents
        End If
        End If
   If Left(cell.Value, 3) = "XDV" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) >= 1 And cell(1, 9).Value = "Ponticello inseribile" Then
         cell(1, 7).ClearContents
        cell(1, 8).ClearContents
        End If
        End If
        
           If Left(cell.Value, 3) = "XDV" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) >= 1 And cell(1, 9).Value = "Insertable jumper" Then
        cell(1, 7).ClearContents
        cell(1, 8).ClearContents
        End If
        End If
   If Left(cell.Value, 3) = "XDV" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) >= 1 And cell(1, 9).Value = "Saddle jumper" Then
        cell(1, 7).ClearContents
        cell(1, 8).ClearContents
        End If
        End If
        
 Next
        
     '------------------XDC -------------------------

    
Set MyPlage = Range("A15:A1000")

    For Each cell In MyPlage

        
           If Left(cell.Value, 3) = "XDC" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) > 1 And cell(1, 9).Value = "Ponticello a staffa" Then
        cell(1, 9).Value = "Ponticello a filo"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
   If Left(cell.Value, 3) = "XDC" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) > 1 And cell(1, 9).Value = "Ponticello inseribile" Then
        cell(1, 9).Value = "Ponticello a filo"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
        
           If Left(cell.Value, 3) = "XDC" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) > 1 And cell(1, 9).Value = "Insertable jumper" Then
        cell(1, 9).Value = "Wire jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
        If Left(cell.Value, 3) = "XDC" And cell.Value = cell(1, 4).Value Then
        If Abs(cell(1, 2).Value - cell(1, 5).Value) > 1 And cell(1, 9).Value = "Saddle jumper" Then
        cell(1, 9).Value = "Wire jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
        

        
         If IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "XDC" And (cell(1, 9).Value = "Wire jumper" Or cell(1, 9).Value = "Ponticello a filo") Then
        XDC = InputBox("Please add cross-section of conductors between" & vbNewLine & cell(1, 3) & " and " & cell(1, 6), "Wire jumper  between " & cell(1, 3) & " and " & cell(1, 6), wire)
        cell(1, 7).Value = XDC
        cell(1, 8).Value = "bk"
        End If
        
   If Left(cell.Value, 3) = "XDC" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) = 1 And IsEmpty(cell(1, 7).Value) Then
    answer = MsgBox("Is posible to have XDC matal jumper?" & vbNewLine & "If this is posible to have XDC matal jumper - Yes", vbYesNo + vbQuestion, "XDC matal jumper")
    If answer = vbNo Then
    XDC = InputBox("Please add cross-section of conductors between" & cell(1, 3) & " and " & cell(1, 6), "Wire jumper  between " & cell(1, 3) & " and " & cell(1, 6), wire)
        If XDC = vbOK Then
        cell(1, 7).Value = XDC
        cell(1, 8).Value = "bk"
        cell(1, 9).Value = "Wire jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        Else
        End If
        End If
        End If


           If IsEmpty(cell(1, 8).Value) And Left(cell.Value, 3) = "XDC" And cell.Value = cell(1, 4).Value Then
        If Abs(cell(1, 2).Value - cell(1, 5).Value) > 1 Then
        cell(1, 8).Value = "bk"
        cell(1, 8).Font.ColorIndex = 3
        cell(1, 8).Font.Bold = True
        End If
        End If

    Next
 

 
 
     '------------------XDM -------------------------

Set MyPlage = Range("A15:A1000")

    For Each cell In MyPlage

        
           If Left(cell.Value, 3) = "XDM" And cell.Value = cell(1, 4).Value And cell(1, 9).Value = "Ponticello a staffa" Then
        cell(1, 9).Value = "Ponticello a filo"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        
            
           If Left(cell.Value, 3) = "XDM" And cell.Value = cell(1, 4).Value And cell(1, 9).Value = "Ponticello inseribile" Then
        cell(1, 9).Value = "Ponticello a filo"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
            

        
           If Left(cell.Value, 3) = "XDM" And cell.Value = cell(1, 4).Value And cell(1, 9).Value = "Insertable jumper" Then
        cell(1, 9).Value = "Wire jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
           If Left(cell.Value, 3) = "XDM" And cell.Value = cell(1, 4).Value And cell(1, 9).Value = "Saddle jumper" Then
        cell(1, 9).Value = "Wire jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        
 Next
 
 '--------------------------------------PGA to PGW------------------------------

Set MyPlage = Range("A15:A1000")

    For Each cell In MyPlage

        
           If Left(cell.Value, 2) = "PG" And cell.Value = cell(1, 4).Value And cell(1, 9).Value = "Ponticello a staffa" Then
        cell(1, 9).Value = "Ponticello a filo"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        
            
           If Left(cell.Value, 2) = "PG" And cell.Value = cell(1, 4).Value And cell(1, 9).Value = "Ponticello inseribile" Then
        cell(1, 9).Value = "Ponticello a filo"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
            

        
           If Left(cell.Value, 2) = "PG" And cell.Value = cell(1, 4).Value And cell(1, 9).Value = "Insertable jumper" Then
        cell(1, 9).Value = "Wire jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
           If Left(cell.Value, 2) = "PG" And cell.Value = cell(1, 4).Value And cell(1, 9).Value = "Saddle jumper" Then
        cell(1, 9).Value = "Wire jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        
 Next
 
 
 '--------------------------------------Selector Switch------------------------------

Set MyPlage = Range("A15:A1000")

    For Each cell In MyPlage

        
           If Left(cell.Value, 2) = "SF" And cell.Value = cell(1, 4).Value And cell(1, 9).Value = "Ponticello a staffa" Then
        cell(1, 9).Value = "Ponticello a filo"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        
            
           If Left(cell.Value, 2) = "SF" And cell.Value = cell(1, 4).Value And cell(1, 9).Value = "Ponticello inseribile" Then
        cell(1, 9).Value = "Ponticello a filo"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
            

        
           If Left(cell.Value, 2) = "SF" And cell.Value = cell(1, 4).Value And cell(1, 9).Value = "Insertable jumper" Then
        cell(1, 9).Value = "Wire jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
           If Left(cell.Value, 2) = "SF" And cell.Value = cell(1, 4).Value And cell(1, 9).Value = "Saddle jumper" Then
        cell(1, 9).Value = "Wire jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        
 Next
 
 '--------------------------------------BT- Thermostat------------------------------

Set MyPlage = Range("A15:A1000")

    For Each cell In MyPlage

        
           If Left(cell.Value, 2) = "BT" And cell.Value = cell(1, 4).Value And cell(1, 9).Value = "Ponticello a staffa" Then
        cell(1, 9).Value = "Ponticello a filo"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        
            
           If Left(cell.Value, 2) = "BT" And cell.Value = cell(1, 4).Value And cell(1, 9).Value = "Ponticello inseribile" Then
        cell(1, 9).Value = "Ponticello a filo"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
            

        
           If Left(cell.Value, 2) = "BT" And cell.Value = cell(1, 4).Value And cell(1, 9).Value = "Insertable jumper" Then
        cell(1, 9).Value = "Wire jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
           If Left(cell.Value, 2) = "BT" And cell.Value = cell(1, 4).Value And cell(1, 9).Value = "Saddle jumper" Then
        cell(1, 9).Value = "Wire jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        
 Next
 
 '--------------------------------------TB------------------------------

Set MyPlage = Range("A15:A1000")

    For Each cell In MyPlage

        
           If Left(cell.Value, 2) = "TB" And cell.Value = cell(1, 4).Value And cell(1, 9).Value = "Ponticello a staffa" Then
        cell(1, 9).Value = "Ponticello a filo"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        
            
           If Left(cell.Value, 2) = "TB" And cell.Value = cell(1, 4).Value And cell(1, 9).Value = "Ponticello inseribile" Then
        cell(1, 9).Value = "Ponticello a filo"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
            

        
           If Left(cell.Value, 2) = "TB" And cell.Value = cell(1, 4).Value And cell(1, 9).Value = "Insertable jumper" Then
        cell(1, 9).Value = "Wire jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
           If Left(cell.Value, 2) = "TB" And cell.Value = cell(1, 4).Value And cell(1, 9).Value = "Saddle jumper" Then
        cell(1, 9).Value = "Wire jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        
 Next

'------------------XDI -------------------------

Set MyPlage = Range("A15:A1000")

    For Each cell In MyPlage


            If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "XDI" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) = 1 And cell(1, 9).Value = "Conductor / wire" Then
           answer = MsgBox("Is it posible to have XDI wire jumper?" & vbNewLine & "If this is posible to have XDI Wire jumper - Yes", vbYesNo + vbQuestion, "XDI jumper")
        If answer = vbNo Then
        cell(1, 7).ClearContents
        cell(1, 8).ClearContents
        cell(1, 9).Value = "Saddle jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        Else
        End If
        End If
                  If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "XDI" And cell.Value = cell(1, 4).Value Then
                  If Abs(cell(1, 2).Value - cell(1, 5).Value) = 1 And cell(1, 9).Value = "Wire jumper" Then
                  answer = MsgBox("Is it posible to have XDI wire jumper?" & vbNewLine & "If this is posible to have XDI Wire jumper - Yes", vbYesNo + vbQuestion, "XDI jumper")
        If answer = vbNo Then
        cell(1, 7).ClearContents
        cell(1, 8).ClearContents
        cell(1, 9).Value = "Saddle jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        Else
        End If
        End If
        
        
        
        
        If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "XDI" And cell.Value = cell(1, 4).Value Then
        If Abs(cell(1, 2).Value - cell(1, 5).Value) = 1 And cell(1, 9).Value = "Conduttore/filo" Then
        cell(1, 7).ClearContents
        cell(1, 8).ClearContents
        cell(1, 9).Value = "Ponticello a staffa"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True

        End If
        End If
               
               If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "XDI" And cell.Value = cell(1, 4).Value Then
        If Abs(cell(1, 2).Value - cell(1, 5).Value) = 1 And cell(1, 9).Value = "Ponticello a filo" Then
        cell(1, 7).ClearContents
        cell(1, 8).ClearContents
        cell(1, 9).Value = "Ponticello a staffa"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        
        End If
        End If
        
           If Left(cell.Value, 3) = "XDI" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) > 1 And cell(1, 9).Value = "Ponticello a staffa" Then
        cell(1, 9).Value = "Ponticello a filo"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
        
        If Left(cell.Value, 3) = "XDI" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) > 1 And cell(1, 9).Value = "Ponticello inseribile" Then
        cell(1, 9).Value = "Ponticello a filo"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
        
           If Left(cell.Value, 3) = "XDI" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) > 1 And cell(1, 9).Value = "Insertable jumper" Then
        cell(1, 9).Value = "Wire jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
         If Left(cell.Value, 3) = "XDI" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) > 1 And cell(1, 9).Value = "Saddle jumper" Then
        cell(1, 9).Value = "Wire jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
        
        
        
        '-------------------clear--------------------------------
        
            If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "XDI" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) = 1 And cell(1, 9).Value = "Ponticello a staffa" Then
        cell(1, 7).ClearContents
        cell(1, 8).ClearContents
        End If
        End If
                  If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "XDI" And cell.Value = cell(1, 4).Value Then
                  If Abs(cell(1, 2).Value - cell(1, 5).Value) = 1 And cell(1, 9).Value = "Ponticello inseribile" Then
        cell(1, 7).ClearContents
        cell(1, 8).ClearContents

        End If
        End If
        
        
        If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "XDI" And cell.Value = cell(1, 4).Value Then
        If Abs(cell(1, 2).Value - cell(1, 5).Value) = 1 And cell(1, 9).Value = "Insertable jumper" Then
         cell(1, 7).ClearContents
        cell(1, 8).ClearContents

        End If
        End If
               
               If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "XDI" And cell.Value = cell(1, 4).Value Then
        If Abs(cell(1, 2).Value - cell(1, 5).Value) = 1 And cell(1, 9).Value = "Saddle jumper" Then
        cell(1, 7).ClearContents
        cell(1, 8).ClearContents
        
        End If
        End If
        
                         If IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "XDI" And (cell(1, 9).Value = "Wire jumper" Or cell(1, 9).Value = "Ponticello a filo") Then
        XDI = InputBox("Please add cross-section of conductors between" & cell(1, 3) & " and " & cell(1, 6), "Wire jumper  between " & cell(1, 3) & " and " & cell(1, 6), wire)
        cell(1, 7).Value = XDI
        cell(1, 8).Value = "bk"
        End If
        
        
        Next
    '--------------------------------------RAR------------------------------
    Set MyPlage = Range("A15:A1000")

    For Each cell In MyPlage

          If Left(cell.Value, 3) = "RAR" Then
          answer = MsgBox("Do you want to clear wire connection between " & vbNewLine & cell(1, 3).Value & " and " & cell(1, 6).Value, vbYesNo + vbQuestion, "RAR section")
          If answer = vbYes Then
            cell(1, 7).ClearContents
            cell(1, 8).ClearContents
         End If
            End If
 Next
                 
                 
 Set MyPlage = Range("D15:D1000")

    For Each cell In MyPlage
    

          If Left(cell.Value, 3) = "RAR" Then
          answer = MsgBox("Do you want to clear wire connection between " & vbNewLine & cell(1, 3).Value & " and " & cell.Offset(1, -1).Value, vbYesNo + vbQuestion, "RAR section")
        If answer = vbYes Then
            cell(1, 4).ClearContents
            cell(1, 5).ClearContents
        End If
        End If
Next

'------------------XDX -------------------------

Set MyPlage = Range("A15:A1000")

    For Each cell In MyPlage


          If Left(cell.Value, 3) = "XDX" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) = 1 And cell(1, 9).Value = "Conductor / wire" Then
        cell(1, 9).Value = "Saddle jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        cell(1, 7).ClearContents
        cell(1, 8).ClearContents
        End If
        End If
                  If Left(cell.Value, 3) = "XDX" And cell.Value = cell(1, 4).Value Then
                  If Abs(cell(1, 2).Value - cell(1, 5).Value) = 1 And cell(1, 9).Value = "Wire jumper" Then
        cell(1, 9).Value = "Saddle jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        cell(1, 7).ClearContents
        cell(1, 8).ClearContents
        End If
        End If
        
        If Left(cell.Value, 3) = "XDX" And cell.Value = cell(1, 4).Value Then
        If Abs(cell(1, 2).Value - cell(1, 5).Value) = 1 And cell(1, 9).Value = "Conduttore/filo" Then
        cell(1, 9).Value = "Ponticello a staffa"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        cell(1, 7).ClearContents
        cell(1, 8).ClearContents
        End If
         End If
               
               If Left(cell.Value, 3) = "XDX" And cell.Value = cell(1, 4).Value Then
        If Abs(cell(1, 2).Value - cell(1, 5).Value) = 1 And cell(1, 9).Value = "Ponticello a filo" Then
        cell(1, 9).Value = "Ponticello a staffa"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        cell(1, 7).ClearContents
        cell(1, 8).ClearContents
        End If
        End If
        
           If Left(cell.Value, 3) = "XDX" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) > 1 And cell(1, 9).Value = "Ponticello a staffa" Then
        cell(1, 9).Value = "Ponticello a filo"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
   If Left(cell.Value, 3) = "XDX" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) > 1 And cell(1, 9).Value = "Ponticello inseribile" Then
        cell(1, 9).Value = "Ponticello a filo"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
        
           If Left(cell.Value, 3) = "XDX" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) > 1 And cell(1, 9).Value = "Insertable jumper" Then
        cell(1, 9).Value = "Wire jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
   If Left(cell.Value, 3) = "XDX" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) > 1 And cell(1, 9).Value = "Saddle jumper" Then
        cell(1, 9).Value = "Wire jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
 
             If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "XDX" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) = 1 And cell(1, 9).Value = "Ponticello a staffa" Then
        cell(1, 7).ClearContents
        cell(1, 8).ClearContents
        End If
        End If
                  If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "XDX" And cell.Value = cell(1, 4).Value Then
                  If Abs(cell(1, 2).Value - cell(1, 5).Value) = 1 And cell(1, 9).Value = "Ponticello inseribile" Then
        cell(1, 7).ClearContents
        cell(1, 8).ClearContents
        End If
        End If

        
        If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "XDX" And cell.Value = cell(1, 4).Value Then
        If Abs(cell(1, 2).Value - cell(1, 5).Value) = 1 And cell(1, 9).Value = "Insertable jumper" Then
         cell(1, 7).ClearContents
        cell(1, 8).ClearContents

        End If
        End If
               
               If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "XDX" And cell.Value = cell(1, 4).Value Then
        If Abs(cell(1, 2).Value - cell(1, 5).Value) = 1 And cell(1, 9).Value = "Saddle jumper" Then
        cell(1, 7).ClearContents
        cell(1, 8).ClearContents

        End If
        End If
        
        
                If IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "XDX" And (cell(1, 9).Value = "Wire jumper" Or cell(1, 9).Value = "Ponticello a filo") Then
        XDX = InputBox("Please add cross-section of conductors between" & vbNewLine & cell(1, 3) & " and " & cell(1, 6), "Wire jumper between " & cell(1, 3) & " and " & cell(1, 6), wire)
        cell(1, 7).Value = XDX
        cell(1, 8).Value = "bk"
        End If
        
        
 
 
  '---------------------------Wire Bridges----------------------------------------------
  
  

   If cell.Value = "XDB1" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) >= 1 And cell(1, 9).Value = "Ponticello a staffa" Then
        cell(1, 9).Value = "Ponticello a filo"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
   If cell.Value = "XDB1" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) >= 1 And cell(1, 9).Value = "Ponticello inseribile" Then
        cell(1, 9).Value = "Ponticello a filo"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
        
           If cell.Value = "XDB1" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) >= 1 And cell(1, 9).Value = "Insertable jumper" Then
        cell(1, 9).Value = "Wire jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
   If cell.Value = "XDB1" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) >= 1 And cell(1, 9).Value = "Saddle jumper" Then
        cell(1, 9).Value = "Wire jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
        
          If IsEmpty(cell(1, 7).Value) And cell.Value = "XDB1" And (cell(1, 9).Value = "Wire jumper" Or cell(1, 9).Value = "Ponticello a filo") Then
        XDB1 = InputBox("Please add cross-section of conductors between" & vbNewLine & cell(1, 3) & " and " & cell(1, 6), "Wire jumper between " & cell(1, 3) & " and " & cell(1, 6), wire)
        cell(1, 7).Value = XDB1
        cell(1, 8).Value = "bk"
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        End If
        
        
           If cell.Value = "XDE" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) >= 1 And cell(1, 9).Value = "Ponticello a staffa" Then
        cell(1, 9).Value = "Ponticello a filo"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
   If cell.Value = "XDE" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) >= 1 And cell(1, 9).Value = "Ponticello inseribile" Then
        cell(1, 9).Value = "Ponticello a filo"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
        
           If cell.Value = "XDE" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) >= 1 And cell(1, 9).Value = "Insertable jumper" Then
        cell(1, 9).Value = "Wire jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
   If cell.Value = "XDE" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) >= 1 And cell(1, 9).Value = "Saddle jumper" Then
        cell(1, 9).Value = "Wire jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
        
        
                    If IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "XDE" And (cell(1, 9).Value = "Wire jumper" Or cell(1, 9).Value = "Ponticello a filo") Then
        XDE = InputBox("Please add cross-section of conductors between" & vbNewLine & cell(1, 3) & " and " & cell(1, 6), "Wire jumper between " & cell(1, 3) & " and " & cell(1, 6), wire)
        cell(1, 7).Value = XDE
        cell(1, 8).Value = "bk"
        End If
              
        
        
           If cell.Value = "XDT" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) >= 1 And cell(1, 9).Value = "Ponticello a staffa" Then
        cell(1, 9).Value = "Ponticello a filo"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
   If cell.Value = "XDT" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) >= 1 And cell(1, 9).Value = "Ponticello inseribile" Then
        cell(1, 9).Value = "Ponticello a filo"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
        
           If cell.Value = "XDT" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) >= 1 And cell(1, 9).Value = "Insertable jumper" Then
        cell(1, 9).Value = "Wire jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
   If cell.Value = "XDT" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) >= 1 And cell(1, 9).Value = "Saddle jumper" Then
        cell(1, 9).Value = "Wire jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
                  If IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "XDT" And (cell(1, 9).Value = "Wire jumper" Or cell(1, 9).Value = "Ponticello a filo") Then
        XDT = InputBox("Please add cross-section of conductors between" & vbNewLine & cell(1, 3) & " and " & cell(1, 6), "Wire jumper between " & cell(1, 3) & " and " & cell(1, 6), wire)
        cell(1, 7).Value = XDT
        cell(1, 8).Value = "bk"
        End If
        
  Next
 
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True


        
 End Sub
  
   


   
