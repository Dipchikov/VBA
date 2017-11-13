Attribute VB_Name = "Legend_of_colours"
Sub Legend_of_colours()

On Error Resume Next

'------------------------------------auto filter--------------------------------------
    ActiveSheet.ShowAllData
    ActiveWorkbook.Worksheets("Wiring table").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Wiring table").AutoFilter.Sort.SortFields.Add Key _
        :=Range("A14:A550"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Wiring table").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    

'------------------CLEAR COLOUR FIRST -------------------------

Range("A15:L1000").Interior.ColorIndex = 0


'------------------Inside Wiring -------------------------


Set MyPlage = Range("A15:A1000")

    For Each cell In MyPlage
    

          If Left(cell.Value, 2) = "BT" Then
        cell(1, 11).Interior.ColorIndex = 40
        End If
        
            If Left(cell.Value, 2) = "BR" Then
        cell(1, 11).Interior.ColorIndex = 40
        End If

         '  If cell.Value = "AA1" Then
        'cell(1, 11).Interior.ColorIndex = 40
        'End If
        
                   'If cell.Value = "AA2" Then
        'cell(1, 11).Interior.ColorIndex = 40
       ' End If
        
                   'If cell.Value = "AA3" Then
        'cell(1, 11).Interior.ColorIndex = 40
        'End If
        
                
                   'If cell.Value = "AA4" Then
        'cell(1, 11).Interior.ColorIndex = 40
        'End If
        
                If cell.Value = "XDB1" Then
        cell(1, 11).Interior.ColorIndex = 40
        End If
                If Left(cell.Value, 3) = "XDE" Then
        cell(1, 11).Interior.ColorIndex = 40
        End If
    
                If Left(cell.Value, 3) = "XDT" Then
        cell(1, 11).Interior.ColorIndex = 40
        End If
                If cell.Value = "PFV1" Then
        cell(1, 11).Interior.ColorIndex = 40
        End If
        
        If Left(cell.Value, 3) = "FCM" Then
            cell(1, 11).Interior.ColorIndex = 40
        End If
                
        If Left(cell.Value, 2) = "TB" Then
        cell(1, 11).Interior.ColorIndex = 40
        End If
        
        If Left(cell.Value, 3) = "XDX" Then
        cell(1, 11).Interior.ColorIndex = 40
        End If

        
        If cell.Value = "XDA" Then
        cell(1, 11).Interior.ColorIndex = 40
        End If

        If cell.Value = "XDV" Then
        cell(1, 11).Interior.ColorIndex = 40
        End If

        If Left(cell.Value, 3) = "XDI" Then
        cell(1, 11).Interior.ColorIndex = 40
        End If

        If Left(cell.Value, 3) = "XDC" Then
        cell(1, 11).Interior.ColorIndex = 40
        End If

        If cell.Value = "K1" Then
        cell(1, 11).Interior.ColorIndex = 40
        End If
        
        If cell.Value = "K2" Then
        cell(1, 11).Interior.ColorIndex = 40
        End If
        
        If cell.Value = "K3" Then
        cell(1, 11).Interior.ColorIndex = 40
        End If
        
        If cell.Value = "K4" Then
        cell(1, 11).Interior.ColorIndex = 40
        End If
                    
             If Left(cell.Value, 3) = "KFA" Then
        cell(1, 11).Interior.ColorIndex = 40
        End If
        If Left(cell.Value, 4) = "KFP." Then
        cell(1, 11).Interior.ColorIndex = 40
        End If
        
            If Left(cell.Value, 3) = "KFC" Then
        cell(1, 11).Interior.ColorIndex = 40
        End If
        
            If Left(cell.Value, 3) = "KFT" Then
        cell(1, 11).Interior.ColorIndex = 40
        End If
        
            If Left(cell.Value, 3) = "KFO" Then
        cell(1, 11).Interior.ColorIndex = 40
        End If
        
        If Left(cell.Value, 3) = "XDC" Then
        cell(1, 11).Interior.ColorIndex = 40
        End If


        If Left(cell.Value, 3) = "TFS" Then
        cell(1, 11).Interior.ColorIndex = 40
        End If
           
        If Left(cell.Value, 3) = "RAR" Then
        cell(1, 11).Interior.ColorIndex = 40
        End If
                   
        If Left(cell.Value, 2) = "XE" Then
        cell(1, 11).Interior.ColorIndex = 40
        End If
        
           If Left(cell.Value, 3) = "XDS" Then
        cell(1, 11).Interior.ColorIndex = 40
        End If
        
    Next
    
      '-----------------Shielded cable--------------------------------
    
    Set MyPlage = Range("H15:H1000")

    For Each cell In MyPlage
    
        If cell.Value = "SH" Then
            cell(1, 4).Interior.ColorIndex = 6
            cell(1, 5).Interior.ColorIndex = 6
        End If
                If cell.Value = "sh" Then
            cell(1, 4).Interior.ColorIndex = 6
            cell(1, 5).Interior.ColorIndex = 6
        End If
        
             If cell.Value = "wh/og" Then
            cell(1, 4).Interior.ColorIndex = 6
            cell(1, 5).Interior.ColorIndex = 6
        End If
        
            If cell.Value = "og/wh" Then
            cell(1, 4).Interior.ColorIndex = 6
            cell(1, 5).Interior.ColorIndex = 6
        End If
    Next
    '----------------------------Door Wireing ----------------------------
    
    
    Set MyPlage = Range("A15:A1000")
        For Each cell In MyPlage
        
           If Left(cell.Value, 3) = "SFA" Then
        cell(1, 11).Interior.ColorIndex = 43
        End If
           If Left(cell.Value, 3) = "SFO" Then
        cell(1, 11).Interior.ColorIndex = 43
        End If
        
         If Left(cell.Value, 3) = "SFM" Then
        cell(1, 11).Interior.ColorIndex = 43
        End If
        
            If Left(cell.Value, 3) = "KFL" Then
        cell(1, 11).Interior.ColorIndex = 43
        End If
        
                    If Left(cell.Value, 3) = "K86" Then
        cell(1, 11).Interior.ColorIndex = 43
        End If
        
           If Left(cell.Value, 3) = "SFU" Then
        cell(1, 11).Interior.ColorIndex = 43
        End If
        
        If Left(cell.Value, 3) = "PFW" Then
            cell(1, 11).Interior.ColorIndex = 43
        End If
                
        If Left(cell.Value, 3) = "PFY" Then
            cell(1, 11).Interior.ColorIndex = 43
        End If
        
            If Left(cell.Value, 3) = "PFS" Then
            cell(1, 11).Interior.ColorIndex = 43
        End If
        
            If Left(cell.Value, 3) = "PFR" Then
            cell(1, 11).Interior.ColorIndex = 43
        End If
        
            If Left(cell.Value, 3) = "SFC" Then
            cell(1, 11).Interior.ColorIndex = 43
        End If
        
        If Left(cell.Value, 3) = "SFS" Then
        cell(1, 11).Interior.ColorIndex = 43
        End If
          
        If Left(cell.Value, 3) = "XDM" Then
        cell(1, 11).Interior.ColorIndex = 43
        End If
        
                
        If Left(cell.Value, 3) = "PFG" Then
        cell(1, 11).Interior.ColorIndex = 43
        End If
        
       If Left(cell.Value, 3) = "PGM" Then
        cell(1, 11).Interior.ColorIndex = 43
        End If
        
               If Left(cell.Value, 3) = "PGC" Then
        cell(1, 11).Interior.ColorIndex = 43
        End If
        
           If Left(cell.Value, 3) = "PGH" Then
        cell(1, 11).Interior.ColorIndex = 43
        End If
        
                If Left(cell.Value, 3) = "PGF" Then
        cell(1, 11).Interior.ColorIndex = 43
        End If
        
        If Left(cell.Value, 3) = "PGA" Then
        cell(1, 11).Interior.ColorIndex = 43
        End If
                
        If Left(cell.Value, 3) = "PGV" Then
        cell(1, 11).Interior.ColorIndex = 43
        End If
                If Left(cell.Value, 3) = "PGI" Then
        cell(1, 11).Interior.ColorIndex = 43
        End If
        
           If Left(cell.Value, 3) = "PFX" Then
        cell(1, 11).Interior.ColorIndex = 43
        End If
        
        
        If Left(cell.Value, 3) = "SFV" Then
        cell(1, 11).Interior.ColorIndex = 43
        End If
        


    Next
 
 '---------------------Wireing - XDB----------------------------
    
    Set MyPlage = Range("D15:D1000")

    For Each cell In MyPlage
    
        If cell.Value = "XDB" Then
            cell(1, 8).Interior.ColorIndex = 37
        End If
        
           If cell.Value = "XDB91" Then
            cell(1, 8).Interior.ColorIndex = 37
        End If
        
           If cell.Value = "XDB10" Then
            cell(1, 8).Interior.ColorIndex = 37
        End If
                   If cell.Value = "XDB89" Then
            cell(1, 8).Interior.ColorIndex = 37
        End If
               If cell.Value = "XDB89" Then
            cell(1, 8).Interior.ColorIndex = 37
        End If

               If cell.Value = "XDB93" Then
            cell(1, 8).Interior.ColorIndex = 37
        End If
               If cell.Value = "XDB95" Then
            cell(1, 8).Interior.ColorIndex = 37
        End If
                       If cell.Value = "XDB96" Then
            cell(1, 8).Interior.ColorIndex = 37
        End If
        
            If cell.Value = "XDB97" Then
            cell(1, 8).Interior.ColorIndex = 37
       End If
    Next
    
      
 
 '---------------------Wireing - 'Ref protection-----------------
 
    Set MyPlage = Range("A15:A1000")

    For Each cell In MyPlage

        If Left(cell.Value, 2) = "AA" Then
            cell(1, 11).Interior.ColorIndex = 44
        End If
        If Left(cell.Value, 3) = "BCR" Then
            cell(1, 11).Interior.ColorIndex = 44
        End If
                If Left(cell.Value, 3) = "BET" Then
            cell(1, 11).Interior.ColorIndex = 44
        End If
    Next
    
    
    

     
        '-------------------------Insertable jumpers"----------------------------------
Set MyPlage = Range("I15:I1000")


    For Each cell In MyPlage
    
        If cell.Value = "Saddle jumper" Then
            cell(1, 1).Interior.ColorIndex = 16
            cell(1, 2).Interior.ColorIndex = 16
            cell(1, 3).Interior.ColorIndex = 16
            cell(1, 4).Interior.ColorIndex = 16
            cell(1, 0).Interior.ColorIndex = 16
            cell(1, -1).Interior.ColorIndex = 16

            
        End If
            If cell.Value = "Insertable jumper" Then
            cell(1, 1).Interior.ColorIndex = 16
            cell(1, 2).Interior.ColorIndex = 16
            cell(1, 3).Interior.ColorIndex = 16
            cell(1, 4).Interior.ColorIndex = 16
            cell(1, 0).Interior.ColorIndex = 16
            cell(1, -1).Interior.ColorIndex = 16

            
        End If

            If cell.Value = "Ponticello a staffa" Then
            cell(1, 1).Interior.ColorIndex = 16
            cell(1, 2).Interior.ColorIndex = 16
            cell(1, 3).Interior.ColorIndex = 16
            cell(1, 4).Interior.ColorIndex = 16
            cell(1, 0).Interior.ColorIndex = 16
            cell(1, -1).Interior.ColorIndex = 16

        End If
            If cell.Value = "Ponticello inseribile" Then
            cell(1, 1).Interior.ColorIndex = 16
            cell(1, 2).Interior.ColorIndex = 16
            cell(1, 3).Interior.ColorIndex = 16
            cell(1, 4).Interior.ColorIndex = 16
            cell(1, 0).Interior.ColorIndex = 16
            cell(1, -1).Interior.ColorIndex = 16

        End If
        
               If cell.Value = "Direct connection" Then
            cell(1, 1).Interior.ColorIndex = 16
            cell(1, 2).Interior.ColorIndex = 16
            cell(1, 3).Interior.ColorIndex = 16
            cell(1, 4).Interior.ColorIndex = 16
            cell(1, 0).Interior.ColorIndex = 16
            cell(1, -1).Interior.ColorIndex = 16

        
        End If
            If cell.Value = "Collegamento diretto" Then
            cell(1, 1).Interior.ColorIndex = 16
            cell(1, 2).Interior.ColorIndex = 16
            cell(1, 3).Interior.ColorIndex = 16
            cell(1, 4).Interior.ColorIndex = 16
            cell(1, 0).Interior.ColorIndex = 16
            cell(1, -1).Interior.ColorIndex = 16

         End If
            Next
            
        
            
              '-------------------------Slap"----------------------------------
             On Error Resume Next
             Set MyPlage = Range("K15:K1000")

            For Each cell In MyPlage
            If cell.Value = "Swap" Then
            cell(1, 1).Interior.ColorIndex = 0
         End If

            
 Next


End Sub
