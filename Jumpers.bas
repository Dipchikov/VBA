Attribute VB_Name = "Jumpers"
Sub Jumpers()

Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False

'---------------------------clear cells BAT-FCF -QAB -BGT -QCE---------------------------------------

Set MyPlage = Range("A15:A1000")

    For Each cell In MyPlage

        If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "BAT" Then
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
            cell(1, 9).Value = "Conduttore/filo"
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
    End If
       If cell.Value <> cell(1, 4).Value And cell(1, 9).Value = "Ponticello inseribile" Then
            cell(1, 9).Value = "Conduttore/filo"
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
    End If
        If cell.Value <> cell(1, 4).Value And cell(1, 9).Value = "Ponticello a filo" Then
            cell(1, 9).Value = "Conduttore/filo"
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
    End If
        
            If cell.Value <> cell(1, 4).Value And cell(1, 9).Value = "Insertable jumper" Then
            cell(1, 9).Value = "Conductor / wire"
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
    End If
      
                 If cell.Value <> cell(1, 4).Value And cell(1, 9).Value = "Saddle jumper" Then
            cell(1, 9).Value = "Conductor / wire"
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
    End If
         
                    If cell.Value <> cell(1, 4).Value And cell(1, 9).Value = "Wire jumper" Then
            cell(1, 9).Value = "Conductor / wire"
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
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
         If Abs(cell(1, 2).Value - cell(1, 5).Value) >= 1 And cell(1, 9).Value = "Ponticello a staffa" Then
        cell(1, 9).Value = "Ponticello a filo"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
   If Left(cell.Value, 3) = "XDC" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) >= 1 And cell(1, 9).Value = "Ponticello inseribile" Then
        cell(1, 9).Value = "Ponticello a filo"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
        
           If Left(cell.Value, 3) = "XDC" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) >= 1 And cell(1, 9).Value = "Insertable jumper" Then
        cell(1, 9).Value = "Wire jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
   If Left(cell.Value, 3) = "XDC" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) >= 1 And cell(1, 9).Value = "Saddle jumper" Then
        cell(1, 9).Value = "Wire jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
         If IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "XDC" Then
        XDC = InputBox("Please add cross-section of conductors ", "Cross-Section of " & cell.Value, "1")
        cell(1, 7).Value = XDC
        cell(1, 8).Value = "bk"
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
        cell(1, 7).ClearContents
        cell(1, 8).ClearContents
        cell(1, 9).Value = "Saddle jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
                  If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "XDI" And cell.Value = cell(1, 4).Value Then
                  If Abs(cell(1, 2).Value - cell(1, 5).Value) = 1 And cell(1, 9).Value = "Wire jumper" Then
        cell(1, 7).ClearContents
        cell(1, 8).ClearContents
         cell(1, 9).Value = "Saddle jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
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
        
        Next
    '--------------------------------------RAR------------------------------
    Set MyPlage = Range("A15:A1000")

    For Each cell In MyPlage

          If Left(cell.Value, 3) = "RAR" Then
          
            cell(1, 7).ClearContents
            cell(1, 8).ClearContents
         End If
 
 Next
                 
                 
 Set MyPlage = Range("D15:D1000")

    For Each cell In MyPlage
    

          If Left(cell.Value, 3) = "RAR" Then
            cell(1, 4).ClearContents
            cell(1, 5).ClearContents
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
 
 
  '---------------------------Wire Bridges----------------------------------------------

   If cell.Value = "XDB1" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) > 1 And cell(1, 9).Value = "Ponticello a staffa" Then
        cell(1, 9).Value = "Ponticello a filo"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
   If cell.Value = "XDB1" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) > 1 And cell(1, 9).Value = "Ponticello inseribile" Then
        cell(1, 9).Value = "Ponticello a filo"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
        
           If cell.Value = "XDB1" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) > 1 And cell(1, 9).Value = "Insertable jumper" Then
        cell(1, 9).Value = "Wire jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
   If cell.Value = "XDB1" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) > 1 And cell(1, 9).Value = "Saddle jumper" Then
        cell(1, 9).Value = "Wire jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
        
           If cell.Value = "XDE" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) > 1 And cell(1, 9).Value = "Ponticello a staffa" Then
        cell(1, 9).Value = "Ponticello a filo"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
   If cell.Value = "XDE" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) > 1 And cell(1, 9).Value = "Ponticello inseribile" Then
        cell(1, 9).Value = "Ponticello a filo"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
        
           If cell.Value = "XDE" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) > 1 And cell(1, 9).Value = "Insertable jumper" Then
        cell(1, 9).Value = "Wire jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
   If cell.Value = "XDE" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) > 1 And cell(1, 9).Value = "Saddle jumper" Then
        cell(1, 9).Value = "Wire jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
           If cell.Value = "XDT" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) > 1 And cell(1, 9).Value = "Ponticello a staffa" Then
        cell(1, 9).Value = "Ponticello a filo"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
   If cell.Value = "XDT" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) > 1 And cell(1, 9).Value = "Ponticello inseribile" Then
        cell(1, 9).Value = "Ponticello a filo"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
        
           If cell.Value = "XDT" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) > 1 And cell(1, 9).Value = "Insertable jumper" Then
        cell(1, 9).Value = "Wire jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
   If cell.Value = "XDT" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) > 1 And cell(1, 9).Value = "Saddle jumper" Then
        cell(1, 9).Value = "Wire jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
  
        
  Next
 
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True


        
 End Sub
  
   


   
