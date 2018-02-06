Attribute VB_Name = "Marine_Legend_of_feruless"
Sub Marine_Legend_of_feruless()
Application.ScreenUpdating = False
Range("T14:T951").Select
Selection.ClearContents

On Error Resume Next



'------------------Inside Wiring -------------------------


Set MyPlage = Range("A14:A1000")

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
           If cell.Value = "AA1" Then
        cell(1, 20).Value = 10
        End If
        
                   If cell.Value = "AA2" Then
        cell(1, 20).Value = 10
       End If
        
                   If cell.Value = "AA3" Then
        cell(1, 20).Value = 10
        End If
        
                
                   If cell.Value = "AA4" Then
        cell(1, 20).Value = 10
        End If
        
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
        
         If Left(cell.Value, 3) = "FCM" And Not (cell(1, 2).Value = 13 Or cell(1, 2).Value = 14 Or cell(1, 2).Value = 21 Or cell(1, 2).Value = 22 Or cell(1, 2).Value = 96 Or cell(1, 2).Value = 95 Or cell(1, 2).Value = 98) Then
            cell(1, 20).Value = 14
        End If
                If Left(cell.Value, 3) = "FCM" And (cell(1, 2).Value = 13 Or cell(1, 2).Value = 14 Or cell(1, 2).Value = 21 Or cell(1, 2).Value = 22 Or cell(1, 2).Value = 96 Or cell(1, 2).Value = 95 Or cell(1, 2).Value = 98) Then
            cell(1, 20).Value = 10
        End If
                
        If Left(cell.Value, 2) = "TB" Then
        cell(1, 20).Value = 10
        End If
        
        If Left(cell.Value, 3) = "XDX" Then
        cell(1, 20).Value = 10
        End If

        
        If cell.Value = "XDA" Then
        cell(1, 20).Value = 14
        End If

        If cell.Value = "XDV" Then
        cell(1, 20).Value = 14
        End If

        If Left(cell.Value, 3) = "XDI" Then
        cell(1, 20).Value = 10
        End If

        If Left(cell.Value, 3) = "XDC" Then
        cell(1, 20).Value = 10
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
    
    
    Set MyPlage = Range("A14:A1000")
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
 
    Set MyPlage = Range("A14:A1000")

    For Each cell In MyPlage

        If Left(cell.Value, 2) = "AA" And Left(cell(1, 2).Value, 5) = "-X130" Then
            cell(1, 20).Value = 14
        End If
        If Left(cell.Value, 2) = "AA" And Not Left(cell(1, 2).Value, 5) = "-X130" And Not cell(1, 7).Value = 4 Then
            cell(1, 20).Value = 10
        End If
            If Left(cell.Value, 2) = "AA" And Not Left(cell(1, 2).Value, 5) = "-X130" And cell(1, 7).Value = 4 Then
            cell(1, 20).Value = 14
        End If
        
        
           If Left(cell.Value, 3) = "BCR" And Not cell(1, 7).Value = 4 Then
            cell(1, 20).Value = 10
        End If
            If Left(cell.Value, 3) = "BCR" And cell(1, 7).Value = 4 Then
            cell(1, 20).Value = 12
        End If
        
                If Left(cell.Value, 3) = "BET" And Not cell(1, 7).Value = 4 Then
            cell(1, 20).Value = 10
        End If
          If Left(cell.Value, 3) = "BET" And cell(1, 7).Value = 4 Then
            cell(1, 20).Value = 12
        End If

             If Left(cell.Value, 3) = "BCP" And Not cell(1, 7).Value = 4 Then
            cell(1, 20).Value = 10
        End If
             If Left(cell.Value, 3) = "BCP" And cell(1, 7).Value = 4 Then
            cell(1, 20).Value = 12
        End If
        
             If Left(cell.Value, 3) = "BCM" And Not cell(1, 7).Value = 4 Then
            cell(1, 20).Value = 10
        End If
                     If Left(cell.Value, 3) = "BCM" And cell(1, 7).Value = 4 Then
            cell(1, 20).Value = 12
        End If
            If Left(cell.Value, 3) = "BCG" And Not cell(1, 7).Value = 4 Then
            cell(1, 20).Value = 10
        End If
             If Left(cell.Value, 3) = "BCG" And cell(1, 7).Value = 4 Then
            cell(1, 20).Value = 12
        End If
        
            If Left(cell.Value, 3) = "BCD" And Not cell(1, 7).Value = 4 Then
            cell(1, 20).Value = 10
        End If
                    If Left(cell.Value, 3) = "BCD" And cell(1, 7).Value = 4 Then
            cell(1, 20).Value = 12
        End If
        
            If Left(cell.Value, 3) = "BCF" And Not cell(1, 7).Value = 4 Then
            cell(1, 20).Value = 10
        End If
                    If Left(cell.Value, 3) = "BCF" And cell(1, 7).Value = 4 Then
            cell(1, 20).Value = 12
        End If
        
            If Left(cell.Value, 3) = "BCP" And Not cell(1, 7).Value = 4 Then
            cell(1, 20).Value = 10
        End If
                    If Left(cell.Value, 3) = "BCP" And cell(1, 7).Value = 4 Then
            cell(1, 20).Value = 12
        End If
        
        
             If Left(cell.Value, 3) = "BCZ" And Not cell(1, 7).Value = 4 Then
            cell(1, 20).Value = 10
        End If
                     If Left(cell.Value, 3) = "BCZ" And cell(1, 7).Value = 4 Then
            cell(1, 20).Value = 12
        End If
        
            If Left(cell.Value, 3) = "BEF" And Not cell(1, 7).Value = 4 Then
            cell(1, 20).Value = 10
        End If
             If Left(cell.Value, 3) = "BEF" And cell(1, 7).Value = 4 Then
            cell(1, 20).Value = 12
        End If
        
             If Left(cell.Value, 3) = "BER" And Not cell(1, 7).Value = 4 Then
            cell(1, 20).Value = 10
        End If
              If Left(cell.Value, 3) = "BER" And cell(1, 7).Value = 4 Then
            cell(1, 20).Value = 12
        End If
             If Left(cell.Value, 3) = "BES" And Not cell(1, 7).Value = 4 Then
            cell(1, 20).Value = 10
        End If
             If Left(cell.Value, 3) = "BES" And cell(1, 7).Value = 4 Then
            cell(1, 20).Value = 12
        End If
            If Left(cell.Value, 3) = "BAR" And Not cell(1, 7).Value = 4 Then
            cell(1, 20).Value = 10
        End If
            If Left(cell.Value, 3) = "BAR" And cell(1, 7).Value = 4 Then
            cell(1, 20).Value = 12
        End If
        
        
        
    Next
    
    
Application.ScreenUpdating = True

End Sub

