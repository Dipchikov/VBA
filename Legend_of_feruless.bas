Attribute VB_Name = "Legend_of_feruless"
Sub Legend_of_feruless()

On Error Resume Next
Set Final = Sheets("Comax")
'------------------CLEAR COLOUR FIRST -------------------------

If Not (Range("L:L10000").Value = "-" Or Range("L:L10000").Value = "Shielded cable") Then


'------------------Inside Wiring -------------------------


Set MyPlage = Range("A15:A1000")

    For Each cell In MyPlage
    

          If Left(cell.Value, 2) = "BT" Then
        cell(1, 20).Value = 10
        End If
        
                  If Left(cell.Value, 2) = "PE" Then
        cell(1, 20).Value = 10
        End If
                  If Left(cell.Value, 2) = "IE" Then
        cell(1, 20).Value = 10
        End If
                  If Left(cell.Value, 2) = "EA" Then
        cell(1, 20).Value = 10
        End If
        
            If Left(cell.Value, 2) = "BR" Then
        cell(1, 20).Value = 10
        End If

          If Left(cell.Value, 2) = "BM" Then
        cell(1, 20).Value = 10
        End If
        
                  If Left(cell.Value, 2) = "BX" Then
        cell(1, 20).Value = 10
        End If
        
                          If Left(cell.Value, 2) = "TS" Then
        cell(1, 20).Value = 10
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
        
                If cell.Value = "XDB1" Then
        cell(1, 20).Value = ""
        End If
                If Left(cell.Value, 3) = "XDE" Then
        cell(1, 20).Value = ""
        End If
        
    
                If Left(cell.Value, 3) = "XDT" Then
        cell(1, 20).Value = ""
        End If
                If Left(cell.Value, 3) = "PFV" Then
        cell(1, 20).Value = 10
        End If
        
            If Left(cell.Value, 3) = "RAD" Then
        cell(1, 20).Value = 10
        End If
        
        If Left(cell.Value, 3) = "FCM" And Not (cell(1, 2).Value = 13 Or cell(1, 2).Value = 14 Or cell(1, 2).Value = 21 Or cell(1, 2).Value = 22) Then
            cell(1, 20).Value = 15
        End If
                If Left(cell.Value, 3) = "FCM" And (cell(1, 2).Value = 13 Or cell(1, 2).Value = 14 Or cell(1, 2).Value = 21 Or cell(1, 2).Value = 22) Then
            cell(1, 20).Value = 10
        End If
                
        If Left(cell.Value, 2) = "TB" Then
        cell(1, 20).Value = 10
        End If
        
        If Left(cell.Value, 3) = "XDX" Then
        cell(1, 20).Value = ""
        End If

        
        If cell.Value = "XDA" Then
        cell(1, 20).Value = 15
        End If

        If cell.Value = "XDV" Then
        cell(1, 20).Value = 15
        End If

        If Left(cell.Value, 3) = "XDI" Then
        cell(1, 20).Value = " "
        End If

        If Left(cell.Value, 3) = "XDC" Then
        cell(1, 20).Value = " "
        End If

        If cell.Value = "K1" Then
        cell(1, 20).Value = 10
        End If
        
        If cell.Value = "K2" Then
        cell(1, 20).Value = 10
        End If
        
        If cell.Value = "K3" Then
        cell(1, 20).Value = 10
        End If
        
        If cell.Value = "K4" Then
        cell(1, 20).Value = 10
        End If
        
        If Left(cell.Value, 2) = "KA" Then
        cell(1, 20).Value = 10
        End If
                    
             If Left(cell.Value, 3) = "KFA" Then
        cell(1, 20).Value = 10
        End If
            If Left(cell.Value, 3) = "RAA" Then
        cell(1, 20).Value = 10
        End If
        
        If Left(cell.Value, 3) = "KFP" Then
        cell(1, 20).Value = 10
        End If
        
          If Left(cell.Value, 3) = "KFE" Then
        cell(1, 20).Value = 10
        End If
        
            If Left(cell.Value, 3) = "KFC" Then
        cell(1, 20).Value = 10
        End If
        
            If Left(cell.Value, 3) = "KFT" Then
        cell(1, 20).Value = 10
        End If
        
            If Left(cell.Value, 3) = "KFO" Then
        cell(1, 20).Value = 10
        End If
        

        If Left(cell.Value, 3) = "TFS" Then
        cell(1, 20).Value = 10
        End If
        
        If Left(cell.Value, 3) = "TFM" Then
        cell(1, 20).Value = 10
        End If
           
        If Left(cell.Value, 3) = "RAR" Then
        cell(1, 20).Value = 10
        End If
                   
        If Left(cell.Value, 2) = "XE" Then
        cell(1, 20).Value = 10
        End If
        
           If Left(cell.Value, 3) = "XDS" Then
        cell(1, 20).Value = 10
        End If
        
    Next
    

    '----------------------------Door Wireing ----------------------------
    
    
    Set MyPlage = Range("A15:A1000")
        For Each cell In MyPlage
        
            If Left(cell.Value, 3) = "SPM" Then
        cell(1, 20).Value = 10
        End If
        
           If Left(cell.Value, 3) = "STF" Then
        cell(1, 20).Value = 10
        End If
        
            If Left(cell.Value, 3) = "SFT" Then
        cell(1, 20).Value = 10
        End If
        
           If Left(cell.Value, 3) = "SFA" Then
        cell(1, 20).Value = 10
        End If
           If Left(cell.Value, 3) = "SFO" Then
        cell(1, 20).Value = 10
        End If
        
         If Left(cell.Value, 3) = "SFM" Then
        cell(1, 20).Value = 10
        End If
        
            If Left(cell.Value, 3) = "KFL" Then
        cell(1, 20).Value = 10
        End If
        
        
           If Left(cell.Value, 3) = "SFU" Then
        cell(1, 20).Value = 10
        End If
        
        If Left(cell.Value, 3) = "PFW" Then
            cell(1, 20).Value = 10
        End If
        
             If Left(cell.Value, 3) = "PGQ" Then
            cell(1, 20).Value = 10
        End If
                
        If Left(cell.Value, 3) = "PFY" Then
            cell(1, 20).Value = 10
        End If
        
            If Left(cell.Value, 3) = "PGW" Then
            cell(1, 20).Value = 10
        End If
        
            If Left(cell.Value, 3) = "PGS" Then
            cell(1, 20).Value = 10
        End If
        
                If Left(cell.Value, 3) = "PFB" Then
            cell(1, 20).Value = 10
        End If
        
            If Left(cell.Value, 3) = "PFS" Then
            cell(1, 20).Value = 10
        End If
        
           If Left(cell.Value, 3) = "PFL" Then
        cell(1, 20).Value = 10
        End If
        
            If Left(cell.Value, 3) = "PFR" Then
            cell(1, 20).Value = 10
        End If
        
            If Left(cell.Value, 3) = "SFC" Then
            cell(1, 20).Value = 10
        End If
        
        If Left(cell.Value, 3) = "SFS" Then
        cell(1, 20).Value = 10
        End If
          
        If Left(cell.Value, 3) = "XDM" Then
        cell(1, 20).Value = 10
        End If
        
                
        If Left(cell.Value, 3) = "PFG" Then
        cell(1, 20).Value = 10
        End If
        
       If Left(cell.Value, 3) = "PGM" Then
        cell(1, 20).Value = 10
        End If
        
               If Left(cell.Value, 3) = "PGC" Then
        cell(1, 20).Value = 10
        End If
        
           If Left(cell.Value, 3) = "PGH" Then
        cell(1, 20).Value = 10
        End If
        
                If Left(cell.Value, 3) = "PGF" Then
        cell(1, 20).Value = 10
        End If
        
        If Left(cell.Value, 3) = "PGA" Then
        cell(1, 20).Value = 10
        End If
                
        If Left(cell.Value, 3) = "PGV" Then
        cell(1, 20).Value = 10
        End If
                If Left(cell.Value, 3) = "PGI" Then
        cell(1, 20).Value = 10
        End If
        
           If Left(cell.Value, 3) = "PFX" Then
        cell(1, 20).Value = 10
        End If
        
        
        If Left(cell.Value, 3) = "SFV" Then
        cell(1, 20).Value = 10
        End If
        
            If Left(cell.Value, 2) = "SF" Then
        cell(1, 20).Value = 10
        End If
        
        '---------------LOCKOUT RELAY---------------------
        
        
                If Left(cell.Value, 3) = "K86" Then
        cell(1, 20).Value = 10
        End If
        


    Next

 '---------------------Wireing - 'Ref protection-----------------
 
    Set MyPlage = Range("A15:A1000")

    For Each cell In MyPlage

        If Left(cell.Value, 2) = "AA" And Left(cell(1, 2).Value, 5) = "-X130" Then
            cell(1, 20).Value = 15
        End If
        If Left(cell.Value, 2) = "AA" And Not Left(cell(1, 2).Value, 5) = "-X130" Then
            cell(1, 20).Value = 10
        End If
        
        
        If Left(cell.Value, 3) = "BCR" Then
            cell(1, 20).Value = 10
        End If
                If Left(cell.Value, 3) = "BET" Then
            cell(1, 20).Value = 10
        End If
             If Left(cell.Value, 3) = "BCP" Then
            cell(1, 20).Value = 10
        End If
             If Left(cell.Value, 3) = "BCM" Then
            cell(1, 20).Value = 10
        End If
            If Left(cell.Value, 3) = "BCG" Then
            cell(1, 20).Value = 10
        End If
            If Left(cell.Value, 3) = "BCD" Then
            cell(1, 20).Value = 10
        End If
            If Left(cell.Value, 3) = "BCF" Then
            cell(1, 20).Value = 10
        End If
            If Left(cell.Value, 3) = "BCP" Then
            cell(1, 20).Value = 10
        End If
             If Left(cell.Value, 3) = "BCZ" Then
            cell(1, 20).Value = 10
        End If
            If Left(cell.Value, 3) = "BEF" Then
            cell(1, 20).Value = 10
        End If
             If Left(cell.Value, 3) = "BER" Then
            cell(1, 20).Value = 10
        End If
             If Left(cell.Value, 3) = "BES" Then
            cell(1, 20).Value = 10
        End If
            If Left(cell.Value, 3) = "BAR" Then
            cell(1, 20).Value = 10
        End If
        
        
        
    Next
    
    
End If

End Sub
