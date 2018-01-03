Attribute VB_Name = "Legend_of_feruless"
Sub Legend_of_feruless()
Application.ScreenUpdating = False
On Error Resume Next
 Range("T14:T951").Select
Selection.ClearContents


'------------------Inside Wiring -------------------------


Set MyPlage = Range("A14:A1000")

    For Each Cell In MyPlage
    

          If Left(Cell.Value, 2) = "BT" Then
        Cell(1, 20).Value = 10
        End If
        
                  If Left(Cell.Value, 2) = "PE" Then
        Cell(1, 20).Value = 10
        End If
                  If Left(Cell.Value, 2) = "IE" Then
        Cell(1, 20).Value = 10
        End If
                  If Left(Cell.Value, 2) = "EA" Then
        Cell(1, 20).Value = 10
        End If
        
            If Left(Cell.Value, 2) = "BR" Then
        Cell(1, 20).Value = 10
        End If

          If Left(Cell.Value, 2) = "BM" Then
        Cell(1, 20).Value = 10
        End If
        
                  If Left(Cell.Value, 2) = "BX" Then
        Cell(1, 20).Value = 10
        End If
        
                          If Left(Cell.Value, 2) = "TS" Then
        Cell(1, 20).Value = 10
        End If
         '  If cell.Value = "AA1" Then
        'cell(1, 20).value = 10
        'End If
        
                   'If cell.Value = "AA2" Then
        'cell(1, 20).value = 10
       ' End If
        
                   'If cell.Value = "AA3" Then
        'cell(1, 20).value = 10
        'End If
        
                
                   'If cell.Value = "AA4" Then
        'cell(1, 20).value = 10
        'End If
        
                If Cell.Value = "XDB1" Then
        Cell(1, 20).Value = ""
        End If
                If Left(Cell.Value, 3) = "XDE" Then
        Cell(1, 20).Value = ""
        End If
        
    
                If Left(Cell.Value, 3) = "XDT" Then
        Cell(1, 20).Value = ""
        End If
                If Left(Cell.Value, 3) = "PFV" Then
        Cell(1, 20).Value = 10
        End If
        
            If Left(Cell.Value, 3) = "RAD" Then
        Cell(1, 20).Value = 10
        End If
        
        If Left(Cell.Value, 3) = "FCM" And Not (Cell(1, 2).Value = 13 Or Cell(1, 2).Value = 14 Or Cell(1, 2).Value = 21 Or Cell(1, 2).Value = 22) Then
            Cell(1, 20).Value = 14
        End If
                If Left(Cell.Value, 3) = "FCM" And (Cell(1, 2).Value = 13 Or Cell(1, 2).Value = 14 Or Cell(1, 2).Value = 21 Or Cell(1, 2).Value = 22) Then
            Cell(1, 20).Value = 10
        End If
                
        If Left(Cell.Value, 2) = "TB" Then
        Cell(1, 20).Value = 10
        End If
        
        If Left(Cell.Value, 3) = "XDX" Then
        Cell(1, 20).Value = ""
        End If

        
        If Cell.Value = "XDA" Then
        Cell(1, 20).Value = 14
        End If

        If Cell.Value = "XDV" Then
        Cell(1, 20).Value = 14
        End If

        If Left(Cell.Value, 3) = "XDI" Then
        Cell(1, 20).Value = ""
        End If

        If Left(Cell.Value, 3) = "XDC" Then
        Cell(1, 20).Value = ""
        End If

        If Cell.Value = "K1" Then
        Cell(1, 20).Value = 10
        End If
        
        If Cell.Value = "K2" Then
        Cell(1, 20).Value = 10
        End If
        
        If Cell.Value = "K3" Then
        Cell(1, 20).Value = 10
        End If
        
        If Cell.Value = "K4" Then
        Cell(1, 20).Value = 10
        End If
        
        If Left(Cell.Value, 2) = "KA" Then
        Cell(1, 20).Value = 10
        End If
                    
             If Left(Cell.Value, 3) = "KFA" Then
        Cell(1, 20).Value = 10
        End If
            If Left(Cell.Value, 3) = "RAA" Then
        Cell(1, 20).Value = 10
        End If
        
        If Left(Cell.Value, 3) = "KFP" Then
        Cell(1, 20).Value = 10
        End If
        
          If Left(Cell.Value, 3) = "KFE" Then
        Cell(1, 20).Value = 10
        End If
        
            If Left(Cell.Value, 3) = "KFC" Then
        Cell(1, 20).Value = 10
        End If
        
            If Left(Cell.Value, 3) = "KFT" Then
        Cell(1, 20).Value = 10
        End If
        
            If Left(Cell.Value, 3) = "KFO" Then
        Cell(1, 20).Value = 10
        End If
        

        If Left(Cell.Value, 3) = "TFS" Then
        Cell(1, 20).Value = 10
        End If
        
        If Left(Cell.Value, 3) = "TFM" Then
        Cell(1, 20).Value = 10
        End If
           
        If Left(Cell.Value, 3) = "RAR" Then
        Cell(1, 20).Value = 10
        End If
                   
        If Left(Cell.Value, 2) = "XE" Then
        Cell(1, 20).Value = 10
        End If
        
           If Left(Cell.Value, 3) = "XDS" Then
        Cell(1, 20).Value = 10
        End If
        
    Next
    

    '----------------------------Door Wireing ----------------------------
    
    
    Set MyPlage = Range("A14:A1000")
        For Each Cell In MyPlage
        
            If Left(Cell.Value, 3) = "SPM" Then
        Cell(1, 20).Value = 10
        End If
        
           If Left(Cell.Value, 3) = "STF" Then
        Cell(1, 20).Value = 10
        End If
        
            If Left(Cell.Value, 3) = "SFT" Then
        Cell(1, 20).Value = 10
        End If
        
           If Left(Cell.Value, 3) = "SFA" Then
        Cell(1, 20).Value = 10
        End If
           If Left(Cell.Value, 3) = "SFO" Then
        Cell(1, 20).Value = 10
        End If
        
         If Left(Cell.Value, 3) = "SFM" Then
        Cell(1, 20).Value = 10
        End If
        
            If Left(Cell.Value, 3) = "KFL" Then
        Cell(1, 20).Value = 10
        End If
        
        
           If Left(Cell.Value, 3) = "SFU" Then
        Cell(1, 20).Value = 10
        End If
        
        If Left(Cell.Value, 3) = "PFW" Then
            Cell(1, 20).Value = 10
        End If
        
             If Left(Cell.Value, 3) = "PGQ" Then
            Cell(1, 20).Value = 10
        End If
                
        If Left(Cell.Value, 3) = "PFY" Then
            Cell(1, 20).Value = 10
        End If
        
            If Left(Cell.Value, 3) = "PGW" Then
            Cell(1, 20).Value = 10
        End If
        
            If Left(Cell.Value, 3) = "PGS" Then
            Cell(1, 20).Value = 10
        End If
        
                If Left(Cell.Value, 3) = "PFB" Then
            Cell(1, 20).Value = 10
        End If
        
            If Left(Cell.Value, 3) = "PFS" Then
            Cell(1, 20).Value = 10
        End If
        
           If Left(Cell.Value, 3) = "PFL" Then
        Cell(1, 20).Value = 10
        End If
        
            If Left(Cell.Value, 3) = "PFR" Then
            Cell(1, 20).Value = 10
        End If
        
            If Left(Cell.Value, 3) = "SFC" Then
            Cell(1, 20).Value = 10
        End If
        
        If Left(Cell.Value, 3) = "SFS" Then
        Cell(1, 20).Value = 10
        End If
          
        If Left(Cell.Value, 3) = "XDM" Then
        Cell(1, 20).Value = 10
        End If
        
                
        If Left(Cell.Value, 3) = "PFG" Then
        Cell(1, 20).Value = 10
        End If
        
       If Left(Cell.Value, 3) = "PGM" Then
        Cell(1, 20).Value = 10
        End If
        
               If Left(Cell.Value, 3) = "PGC" Then
        Cell(1, 20).Value = 10
        End If
        
           If Left(Cell.Value, 3) = "PGH" Then
        Cell(1, 20).Value = 10
        End If
        
                If Left(Cell.Value, 3) = "PGF" Then
        Cell(1, 20).Value = 10
        End If
        
        If Left(Cell.Value, 3) = "PGA" Then
        Cell(1, 20).Value = 10
        End If
                
        If Left(Cell.Value, 3) = "PGV" Then
        Cell(1, 20).Value = 10
        End If
                If Left(Cell.Value, 3) = "PGI" Then
        Cell(1, 20).Value = 10
        End If
        
           If Left(Cell.Value, 3) = "PFX" Then
        Cell(1, 20).Value = 10
        End If
        
        
        If Left(Cell.Value, 3) = "SFV" Then
        Cell(1, 20).Value = 10
        End If
        
            If Left(Cell.Value, 2) = "SF" Then
        Cell(1, 20).Value = 10
        End If
        
        '---------------LOCKOUT RELAY---------------------
        
        
                If Left(Cell.Value, 3) = "K86" Then
        Cell(1, 20).Value = 10
        End If
        


    Next

 '---------------------Wireing - 'Ref protection-----------------
 
    Set MyPlage = Range("A14:A1000")

    For Each Cell In MyPlage

        If Left(Cell.Value, 2) = "AA" And Left(Cell(1, 2).Value, 5) = "-X130" Then
            Cell(1, 20).Value = 14
        End If
        If Left(Cell.Value, 2) = "AA" And Not Left(Cell(1, 2).Value, 5) = "-X130" Then
            Cell(1, 20).Value = 10
        End If
        
        
        If Left(Cell.Value, 3) = "BCR" Then
            Cell(1, 20).Value = 10
        End If
                If Left(Cell.Value, 3) = "BET" Then
            Cell(1, 20).Value = 10
        End If
             If Left(Cell.Value, 3) = "BCP" Then
            Cell(1, 20).Value = 10
        End If
             If Left(Cell.Value, 3) = "BCM" Then
            Cell(1, 20).Value = 10
        End If
            If Left(Cell.Value, 3) = "BCG" Then
            Cell(1, 20).Value = 10
        End If
            If Left(Cell.Value, 3) = "BCD" Then
            Cell(1, 20).Value = 10
        End If
            If Left(Cell.Value, 3) = "BCF" Then
            Cell(1, 20).Value = 10
        End If
            If Left(Cell.Value, 3) = "BCP" Then
            Cell(1, 20).Value = 10
        End If
             If Left(Cell.Value, 3) = "BCZ" Then
            Cell(1, 20).Value = 10
        End If
            If Left(Cell.Value, 3) = "BEF" Then
            Cell(1, 20).Value = 10
        End If
             If Left(Cell.Value, 3) = "BER" Then
            Cell(1, 20).Value = 10
        End If
             If Left(Cell.Value, 3) = "BES" Then
            Cell(1, 20).Value = 10
        End If
            If Left(Cell.Value, 3) = "BAR" Then
            Cell(1, 20).Value = 10
        End If
        
        
        
    Next
    
Application.ScreenUpdating = True

End Sub
