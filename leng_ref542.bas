Attribute VB_Name = "leng_ref542"
Sub leng_ref542()

Dim lr As Long
lr = Range("A" & Rows.Count).End(xlUp).Row
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False

On Error Resume Next


If Error_menu.CheckBox4.Value = True Then


  '----------------------------Door Wireing ----------------------------
    Set MyPlage = Range("A15:A" & lr)
        For Each cell In MyPlage
        
            If Left(cell.Value, 2) = "AA" And Left(cell(1, 4).Value, 3) = "SPM" Then
        cell(1, 11).Value = cell(1, 11).Value + 1000
        End If
        
           If Left(cell.Value, 2) = "AA" And Left(cell(1, 4).Value, 3) = "STF" Then
        cell(1, 11).Value = cell(1, 11).Value + 1000
        End If
        
            If Left(cell.Value, 2) = "AA" And Left(cell(1, 4).Value, 3) = "SFT" Then
        cell(1, 11).Value = cell(1, 11).Value + 1000
        End If
        
           If Left(cell.Value, 2) = "AA" And Left(cell(1, 4).Value, 3) = "SFA" Then
        cell(1, 11).Value = cell(1, 11).Value + 1000
        End If
           If Left(cell.Value, 2) = "AA" And Left(cell(1, 4).Value, 3) = "SFO" Then
        cell(1, 11).Value = cell(1, 11).Value + 1000
        End If
        
         If Left(cell.Value, 2) = "AA" And Left(cell(1, 4).Value, 3) = "SFM" Then
        cell(1, 11).Value = cell(1, 11).Value + 1000
        End If
        
            If Left(cell.Value, 2) = "AA" And Left(cell(1, 4).Value, 3) = "KFL" Then
        cell(1, 11).Value = cell(1, 11).Value + 1000
        End If
        
        
           If Left(cell.Value, 2) = "AA" And Left(cell(1, 4).Value, 3) = "SFU" Then
        cell(1, 11).Value = cell(1, 11).Value + 1000
        End If
        
        If Left(cell.Value, 2) = "AA" And Left(cell(1, 4).Value, 3) = "PFW" Then
            cell(1, 11).Value = cell(1, 11).Value + 1000
        End If
        
             If Left(cell.Value, 2) = "AA" And Left(cell(1, 4).Value, 3) = "PGQ" Then
            cell(1, 11).Value = cell(1, 11).Value + 1000
        End If
                
        If Left(cell.Value, 2) = "AA" And Left(cell(1, 4).Value, 3) = "PFY" Then
            cell(1, 11).Value = cell(1, 11).Value + 1000
        End If
        
            If Left(cell.Value, 2) = "AA" And Left(cell(1, 4).Value, 3) = "PGW" Then
            cell(1, 11).Value = cell(1, 11).Value + 1000
        End If
        
            If Left(cell.Value, 2) = "AA" And Left(cell(1, 4).Value, 3) = "PGS" Then
            cell(1, 11).Value = cell(1, 11).Value + 1000
        End If
        
                If Left(cell.Value, 2) = "AA" And Left(cell(1, 4).Value, 3) = "PFB" Then
            cell(1, 11).Value = cell(1, 11).Value + 1000
        End If
        
            If Left(cell.Value, 2) = "AA" And Left(cell(1, 4).Value, 3) = "PFS" Then
            cell(1, 11).Value = cell(1, 11).Value + 1000
        End If
        
           If Left(cell.Value, 2) = "AA" And Left(cell(1, 4).Value, 3) = "PFL" Then
        cell(1, 11).Value = cell(1, 11).Value + 1000
        End If
        
                   If Left(cell.Value, 2) = "AA" And Left(cell(1, 4).Value, 3) = "PFF" Then
        cell(1, 11).Value = cell(1, 11).Value + 1000
        End If
        
            If Left(cell.Value, 2) = "AA" And Left(cell(1, 4).Value, 3) = "PFR" Then
            cell(1, 11).Value = cell(1, 11).Value + 1000
        End If
        
            If Left(cell.Value, 2) = "AA" And Left(cell(1, 4).Value, 3) = "SFC" Then
            cell(1, 11).Value = cell(1, 11).Value + 1000
        End If
        
        If Left(cell.Value, 2) = "AA" And Left(cell(1, 4).Value, 3) = "SFS" Then
        cell(1, 11).Value = cell(1, 11).Value + 1000
        End If
          
        If Left(cell.Value, 2) = "AA" And Left(cell(1, 4).Value, 3) = "XDM" Then
        cell(1, 11).Value = cell(1, 11).Value + 1000
        End If
        
                
        If Left(cell.Value, 2) = "AA" And Left(cell(1, 4).Value, 3) = "PFG" Then
        cell(1, 11).Value = cell(1, 11).Value + 1000
        End If
        
       If Left(cell.Value, 2) = "AA" And Left(cell(1, 4).Value, 3) = "PGM" Then
        cell(1, 11).Value = cell(1, 11).Value + 1000
        End If
        
               If Left(cell.Value, 2) = "AA" And Left(cell(1, 4).Value, 3) = "PGC" Then
        cell(1, 11).Value = cell(1, 11).Value + 1000
        End If
        
           If Left(cell.Value, 2) = "AA" And Left(cell(1, 4).Value, 3) = "PGH" Then
        cell(1, 11).Value = cell(1, 11).Value + 1000
        End If
        
                If Left(cell.Value, 2) = "AA" And Left(cell(1, 4).Value, 3) = "PGF" Then
        cell(1, 11).Value = cell(1, 11).Value + 1000
        End If
        
        If Left(cell.Value, 2) = "AA" And Left(cell(1, 4).Value, 3) = "PGA" Then
        cell(1, 11).Value = cell(1, 11).Value + 1000
        End If
                
        If Left(cell.Value, 2) = "AA" And Left(cell(1, 4).Value, 3) = "PGV" Then
        cell(1, 11).Value = cell(1, 11).Value + 1000
        End If
                If Left(cell.Value, 2) = "AA" And Left(cell(1, 4).Value, 3) = "PGI" Then
        cell(1, 11).Value = cell(1, 11).Value + 1000
        End If
        
           If Left(cell.Value, 2) = "AA" And Left(cell(1, 4).Value, 3) = "PFX" Then
        cell(1, 11).Value = cell(1, 11).Value + 1000
        End If
        
        
        If Left(cell.Value, 2) = "AA" And Left(cell(1, 4).Value, 3) = "SFV" Then
        cell(1, 11).Value = cell(1, 11).Value + 1000
        End If
        

    
        
        '---------------LOCKOUT RELAY---------------------
        
        
                If Left(cell.Value, 2) = "AA" And Left(cell(1, 4).Value, 3) = "K86" Then
        cell(1, 11).Value = cell(1, 11).Value + 1000
        End If
        
        '------------------Inside Wiring -------------------------

            If Left(cell.Value, 2) = "AA" And Left(cell(1, 4).Value, 2) = "XD" Then
        cell(1, 11).Value = cell(1, 11).Value - 700
        End If
        

          If Left(cell.Value, 2) = "AA" And Left(cell.Value, 2) = "BT" Then
        cell(1, 11).Value = cell(1, 11).Value - 700
        End If
        
                  If Left(cell.Value, 2) = "AA" And cell.Value = "F1" Then
        cell(1, 11).Value = cell(1, 11).Value - 700
        End If
        
                  If Left(cell.Value, 2) = "AA" And Left(cell.Value, 2) = "KM" Then
        cell(1, 11).Value = cell(1, 11).Value - 700
        End If
        
        
         If Left(cell.Value, 2) = "AA" And Left(cell.Value, 2) = "PJ" Then
        cell(1, 11).Value = cell(1, 11).Value - 700
        End If
        
                  If Left(cell.Value, 2) = "AA" And Left(cell.Value, 2) = "PE" Then
        cell(1, 11).Value = cell(1, 11).Value - 700
        End If
                  If Left(cell.Value, 2) = "AA" And Left(cell.Value, 2) = "IE" Then
        cell(1, 11).Value = cell(1, 11).Value - 700
        End If
                  If Left(cell.Value, 2) = "AA" And Left(cell.Value, 2) = "EA" Then
        cell(1, 11).Value = cell(1, 11).Value - 700
        End If
        
            If Left(cell.Value, 2) = "AA" And Left(cell.Value, 2) = "BR" Then
        cell(1, 11).Value = cell(1, 11).Value - 700
        End If

          If Left(cell.Value, 2) = "AA" And Left(cell.Value, 2) = "BM" Then
        cell(1, 11).Value = cell(1, 11).Value - 700
        End If
        
                  If Left(cell.Value, 2) = "AA" And Left(cell.Value, 2) = "BX" Then
        cell(1, 11).Value = cell(1, 11).Value - 700
        End If
        
          If Left(cell.Value, 2) = "AA" And Left(cell.Value, 2) = "TS" Then
        cell(1, 11).Value = cell(1, 11).Value - 700
        End If
    

                If Left(cell.Value, 2) = "AA" And Left(cell(1, 4).Value, 3) = "PFV" Then
        cell(1, 11).Value = cell(1, 11).Value - 700
        End If
        
            If Left(cell.Value, 2) = "AA" And Left(cell(1, 4).Value, 3) = "RAD" Then
        cell(1, 11).Value = cell(1, 11).Value - 700
        End If
        
        If Left(cell.Value, 2) = "AA" And Left(cell(1, 4).Value, 3) = "FCM" Then
            cell(1, 11).Value = cell(1, 11).Value - 700
        End If
                
        If Left(cell.Value, 2) = "AA" And Left(cell.Value, 2) = "TB" Then
        cell(1, 11).Value = cell(1, 11).Value - 700
        End If
        

         If Left(cell.Value, 2) = "AA" And cell.Value = "K1" Then
        cell(1, 11).Value = cell(1, 11).Value - 700
        End If
        
        If Left(cell.Value, 2) = "AA" And cell.Value = "K2" Then
        cell(1, 11).Value = cell(1, 11).Value - 700
        End If
        
        If Left(cell.Value, 2) = "AA" And cell.Value = "K3" Then
        cell(1, 11).Value = cell(1, 11).Value - 700
        End If
        
        If Left(cell.Value, 2) = "AA" And cell.Value = "K4" Then
        cell(1, 11).Value = cell(1, 11).Value - 700
        End If
        
        If Left(cell.Value, 2) = "AA" And Left(cell.Value, 2) = "KA" Then
        cell(1, 11).Value = cell(1, 11).Value - 700
        End If
                    
             If Left(cell.Value, 2) = "AA" And Left(cell(1, 4).Value, 3) = "KFA" Then
        cell(1, 11).Value = cell(1, 11).Value - 700
        End If
            If Left(cell.Value, 2) = "AA" And Left(cell(1, 4).Value, 3) = "RAA" Then
        cell(1, 11).Value = cell(1, 11).Value - 700
        End If
        
        If Left(cell.Value, 2) = "AA" And Left(cell(1, 4).Value, 3) = "KFP" Then
        cell(1, 11).Value = cell(1, 11).Value - 700
        End If
        
          If Left(cell.Value, 2) = "AA" And Left(cell(1, 4).Value, 3) = "KFE" Then
        cell(1, 11).Value = cell(1, 11).Value - 700
        End If
        
            If Left(cell.Value, 2) = "AA" And Left(cell(1, 4).Value, 3) = "KFC" Then
        cell(1, 11).Value = cell(1, 11).Value - 700
        End If
        
            If Left(cell.Value, 2) = "AA" And Left(cell(1, 4).Value, 3) = "KFT" Then
        cell(1, 11).Value = cell(1, 11).Value - 700
        End If
        
            If Left(cell.Value, 2) = "AA" And Left(cell(1, 4).Value, 3) = "KFO" Then
        cell(1, 11).Value = cell(1, 11).Value - 700
        End If
        
        If Left(cell.Value, 2) = "AA" And Left(cell(1, 4).Value, 3) = "XDC" Then
        cell(1, 11).Value = cell(1, 11).Value - 700
        End If


        If Left(cell.Value, 2) = "AA" And Left(cell(1, 4).Value, 3) = "TFS" Then
        cell(1, 11).Value = cell(1, 11).Value - 700
        End If
        
        If Left(cell.Value, 2) = "AA" And Left(cell(1, 4).Value, 3) = "TFM" Then
        cell(1, 11).Value = cell(1, 11).Value - 700
        End If
           
        If Left(cell.Value, 2) = "AA" And Left(cell(1, 4).Value, 3) = "RAR" Then
        cell(1, 11).Value = cell(1, 11).Value - 700
        End If
                   
        If Left(cell.Value, 2) = "AA" And Left(cell.Value, 2) = "XE" Then
        cell(1, 11).Value = cell(1, 11).Value - 700
        End If
        
           If Left(cell.Value, 2) = "AA" And Left(cell(1, 4).Value, 3) = "XDS" Then
        cell(1, 11).Value = cell(1, 11).Value - 700
        End If
        
                 If Left(cell.Value, 2) = "AA" And Left(cell(1, 4).Value, 3) = "TFC" Then
        cell(1, 11).Value = cell(1, 11).Value - 700
        End If
        
        
             If Left(cell.Value, 2) = "AA" And Left(cell(1, 4).Value, 3) = "KLA" Then
        cell(1, 11).Value = cell(1, 11).Value - 700
        End If
        
              If Left(cell.Value, 2) = "AA" And Left(cell(1, 4).Value, 3) = "KLT" Then
        cell(1, 11).Value = cell(1, 11).Value - 700
        End If
        
        If Left(cell.Value, 2) = "AA" And Left(cell(1, 4).Value, 3) = "QBM" Then
        cell(1, 11).Value = cell(1, 11).Value - 700
        End If
    


    Next
    End If
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub
