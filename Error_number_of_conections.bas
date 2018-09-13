Attribute VB_Name = "Error_number_of_conections"
Sub Error_number_of_conections()
Dim lr As Long
lr = Range("A" & Rows.Count).End(xlUp).Row
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
'---------------------------------------------XDC------------------------------
Set MyPlage = Range("A15:A" & lr)
  For Each cell In MyPlage

                    If Left(cell.Value, 3) = "XDC" And cell(1, 13).Value > 2 And cell(1, 13).Value <= 3 Then
        cell(1, 2).Interior.ColorIndex = 46
        End If

                    If Left(cell.Value, 3) = "XDC" And cell(1, 13).Value > 3 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
                   If Left(cell.Value, 3) = "XDC" And cell(1, 13).Value < 2 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If

Next
Set MyPlage = Range("D15:D" & lr)
  For Each cell In MyPlage
        If Left(cell.Value, 3) = "XDC" And cell(1, 11).Value > 2 And cell(1, 11).Value <= 3 Then
        cell(1, 2).Interior.ColorIndex = 46
        End If
                    If Left(cell.Value, 3) = "XDC" And cell(1, 11).Value > 3 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
            If Left(cell.Value, 3) = "XDC" And cell(1, 11).Value < 2 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If
    
Next

'---------------------------------------------XDI------------------------------
Set MyPlage = Range("A15:A" & lr)
  For Each cell In MyPlage


        
                   If Left(cell.Value, 3) = "XDI" And cell(1, 13).Value > 2 And cell(1, 13).Value <= 3 Then
        cell(1, 2).Interior.ColorIndex = 45
        End If
                   If Left(cell.Value, 3) = "XDI" And cell(1, 13).Value > 3 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
        
                           If Left(cell.Value, 3) = "XDI" And cell(1, 13).Value <= 2 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If
        
        
Next


Set MyPlage = Range("D15:D" & lr)
  For Each cell In MyPlage

        
                   If Left(cell.Value, 3) = "XDI" And cell(1, 11).Value > 2 And cell(1, 11).Value <= 3 Then
        cell(1, 2).Interior.ColorIndex = 45
        End If
                   If Left(cell.Value, 3) = "XDI" And cell(1, 11).Value > 3 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
               If Left(cell.Value, 3) = "XDI" And cell(1, 11).Value <= 2 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If
Next

'---------------------------------------------XDX------------------------------
Set MyPlage = Range("A15:A" & lr)
  For Each cell In MyPlage

        
                   If Left(cell.Value, 3) = "XDX" And cell(1, 13).Value >= 5 And cell(1, 13).Value <= 6 Then
        cell(1, 2).Interior.ColorIndex = 45
        End If
            If Left(cell.Value, 3) = "XDX" And cell(1, 13).Value > 6 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
              If Left(cell.Value, 3) = "XDX" And cell(1, 2).Value = 1 And cell(1, 13).Value > 5 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
              If Left(cell.Value, 3) = "XDX" And cell(1, 13).Value <= 4 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If
        
Next
Set MyPlage = Range("D15:D" & lr)
  For Each cell In MyPlage

        
                   If Left(cell.Value, 3) = "XDX" And cell(1, 11).Value >= 5 And cell(1, 11).Value <= 6 Then
        cell(1, 2).Interior.ColorIndex = 45
        End If
                  If Left(cell.Value, 3) = "XDX" And cell(1, 11).Value > 6 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
                 If Left(cell.Value, 3) = "XDX" And cell(1, 2).Value = 1 And cell(1, 11).Value > 5 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
                          If Left(cell.Value, 3) = "XDX" And cell(1, 11).Value <= 4 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If
Next

'---------------------------------------------FCM------------------------------
Set MyPlage = Range("A15:A" & lr)
  For Each cell In MyPlage

        
                   If Left(cell.Value, 3) = "FCM" And cell(1, 13).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
                          
Next

Set MyPlage = Range("D15:D" & lr)
  For Each cell In MyPlage

        
           If Left(cell.Value, 3) = "FCM" And cell(1, 11).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
        
Next
      
'---------------------------------------------KFA to KFZ------------------------------
Set MyPlage = Range("A15:A" & lr)
  For Each cell In MyPlage

        
                   If Left(cell.Value, 2) = "KF" And cell(1, 13).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If

Next

Set MyPlage = Range("D15:D" & lr)
  For Each cell In MyPlage

        
           If Left(cell.Value, 2) = "KF" And cell(1, 11).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        

  Next


'---------------------------------------------Lamps------------------------------
Set MyPlage = Range("A15:A" & lr)
  For Each cell In MyPlage

        
                   If Left(cell.Value, 3) = "PFB" And cell(1, 13).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
                
                   If Left(cell.Value, 3) = "PFG" And cell(1, 13).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
                       If Left(cell.Value, 3) = "PFR" And cell(1, 13).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
             If Left(cell.Value, 3) = "PFY" And cell(1, 13).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
        If Left(cell.Value, 3) = "SPM" And cell(1, 13).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
        If Left(cell.Value, 3) = "SFT" And cell(1, 13).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
        If Left(cell.Value, 3) = "STF" And cell(1, 13).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If

        If Left(cell.Value, 3) = "PFL" And cell(1, 13).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
Next

Set MyPlage = Range("D15:D" & lr)
  For Each cell In MyPlage

        
           If Left(cell.Value, 3) = "PFB" And cell(1, 11).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
           If Left(cell.Value, 3) = "PFG" And cell(1, 11).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
              If Left(cell.Value, 3) = "PFR" And cell(1, 11).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
                      If Left(cell.Value, 3) = "PFY" And cell(1, 11).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
            
     
                If Left(cell.Value, 3) = "SPM" And cell(1, 11).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
              If Left(cell.Value, 3) = "SFT" And cell(1, 11).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
             If Left(cell.Value, 3) = "STF" And cell(1, 11).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
             
             If Left(cell.Value, 3) = "PFL" And cell(1, 11).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
     
Next

'--------------------------------------Selector Switch------------------------------
Set MyPlage = Range("A15:A" & lr)
  For Each cell In MyPlage

        If Left(cell.Value, 2) = "SF" And cell(1, 13).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
Next

Set MyPlage = Range("D15:D" & lr)
  For Each cell In MyPlage

        
         If Left(cell.Value, 2) = "SF" And cell(1, 11).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
Next

'--------------------------------------BT- Thermostat------------------------------
Set MyPlage = Range("A15:A" & lr)
  For Each cell In MyPlage

        If Left(cell.Value, 2) = "BT" And cell(1, 13).Value > 1 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
Next

Set MyPlage = Range("D15:D" & lr)
  For Each cell In MyPlage

        
         If Left(cell.Value, 2) = "BT" And cell(1, 11).Value > 1 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
Next
'--------------------------------------XDM------------------------------
Set MyPlage = Range("A15:A" & lr)
  For Each cell In MyPlage

        If Left(cell.Value, 3) = "XDM" And cell(1, 13).Value > 1 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
Next

Set MyPlage = Range("D15:D" & lr)
  For Each cell In MyPlage

        
         If Left(cell.Value, 3) = "XDM" And cell(1, 11).Value > 1 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
Next
'--------------------------------------PFV------------------------------
Set MyPlage = Range("A15:A" & lr)
  For Each cell In MyPlage

        If Left(cell.Value, 3) = "PFV" And cell(1, 13).Value > 1 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
Next

Set MyPlage = Range("D15:D" & lr)
  For Each cell In MyPlage

        
         If Left(cell.Value, 3) = "PFV" And cell(1, 11).Value > 1 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
    Next
'--------------------------------------PGA to PGW------------------------------
Set MyPlage = Range("A15:A" & lr)
  For Each cell In MyPlage

        If Left(cell.Value, 2) = "PG" And cell(1, 13).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
Next

Set MyPlage = Range("D15:D" & lr)
  For Each cell In MyPlage

        
         If Left(cell.Value, 2) = "PG" And cell(1, 11).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
Next

'---------------------------------------------PGM------------------------------
Set MyPlage = Range("A15:A" & lr)
  For Each cell In MyPlage

        
                   If Left(cell.Value, 3) = "PGM" And cell(1, 13).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
                          
Next

Set MyPlage = Range("D15:D" & lr)
  For Each cell In MyPlage

        
           If Left(cell.Value, 3) = "PGM" And cell(1, 11).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
        
Next
             
 '--------------------------------------RAR------------------------------
Set MyPlage = Range("A15:A" & lr)
  For Each cell In MyPlage

        If Left(cell.Value, 3) = "RAR" And cell(1, 13).Value > 3 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
Next

Set MyPlage = Range("D15:D" & lr)
  For Each cell In MyPlage

        
         If Left(cell.Value, 2) = "RAR" And cell(1, 11).Value > 3 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
Next


'---------------------------REF542---------------------------------------------
If Error_menu.CheckBox4.Value = True Then

Set MyPlage = Range("A15:A" & lr)
  For Each cell In MyPlage


                    If Left(cell.Value, 2) = "AA" And cell(1, 13).Value > 1 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
                    If Left(cell.Value, 2) = "AA" And cell(1, 13).Value <= 1 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If


Next

Set MyPlage = Range("D15:D" & lr)
  For Each cell In MyPlage


                    If Left(cell.Value, 2) = "AA" And cell(1, 11).Value > 1 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
                    If Left(cell.Value, 2) = "AA" And cell(1, 11).Value <= 1 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If


Next
End If

    '---------------------------REF errors---------------------------------------------
    
    If Error_menu.CheckBox4.Value = False Then
    Set MyPlage = Range("A15:A" & lr)

'--------------------------Ref---------------------------------------------
  For Each cell In MyPlage

                    If Left(cell.Value, 2) = "AA" And cell(1, 13).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
                    If Left(cell.Value, 2) = "AA" And cell(1, 13).Value <= 2 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If
        
              If Left(cell.Value, 2) = "AA" And cell(1, 13).Value >= 2 Then
        cell(1, 7).Interior.ColorIndex = 46
        End If
Next

Set MyPlage = Range("D15:D" & lr)
  For Each cell In MyPlage

                    If Left(cell.Value, 2) = "AA" And cell(1, 11).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
                    If Left(cell.Value, 2) = "AA" And cell(1, 11).Value <= 2 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If
        
        
                      If Left(cell.Value, 2) = "AA" And cell(1, 11).Value >= 2 Then
        cell(1, 4).Interior.ColorIndex = 46
        End If
Next
'--------------------------BCR---------------------------------------------
Set MyPlage = Range("A15:A" & lr)

  For Each cell In MyPlage

                    If Left(cell.Value, 3) = "BCR" And cell(1, 13).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
                    If Left(cell.Value, 3) = "BCR" And cell(1, 13).Value <= 2 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If
        
              If Left(cell.Value, 3) = "BCR" And cell(1, 13).Value >= 2 Then
        cell(1, 7).Interior.ColorIndex = 46
        End If
Next

Set MyPlage = Range("D15:D" & lr)
  For Each cell In MyPlage

                    If Left(cell.Value, 3) = "BCR" And cell(1, 11).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
                    If Left(cell.Value, 3) = "BCR" And cell(1, 11).Value <= 2 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If
        
        
                      If Left(cell.Value, 3) = "BCR" And cell(1, 11).Value >= 2 Then
        cell(1, 4).Interior.ColorIndex = 46
        End If
Next
'--------------------------BET---------------------------------------------
Set MyPlage = Range("A15:A" & lr)

  For Each cell In MyPlage

                    If Left(cell.Value, 3) = "BET" And cell(1, 13).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
                    If Left(cell.Value, 3) = "BET" And cell(1, 13).Value <= 2 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If
        
              If Left(cell.Value, 3) = "BET" And cell(1, 13).Value >= 2 Then
        cell(1, 7).Interior.ColorIndex = 46
        End If
Next

Set MyPlage = Range("D15:D" & lr)
  For Each cell In MyPlage

                    If Left(cell.Value, 3) = "BET" And cell(1, 11).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
                    If Left(cell.Value, 3) = "BET" And cell(1, 11).Value <= 2 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If
        
        
                      If Left(cell.Value, 3) = "BET" And cell(1, 11).Value >= 2 Then
        cell(1, 4).Interior.ColorIndex = 46
        End If
        
Next
'--------------------------BCP---------------------------------------------
Set MyPlage = Range("A15:A" & lr)

  For Each cell In MyPlage

                    If Left(cell.Value, 3) = "BCP" And cell(1, 13).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
                    If Left(cell.Value, 3) = "BCP" And cell(1, 13).Value <= 2 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If
        
              If Left(cell.Value, 3) = "BCP" And cell(1, 13).Value >= 2 Then
        cell(1, 7).Interior.ColorIndex = 46
        End If
Next

Set MyPlage = Range("D15:D" & lr)
  For Each cell In MyPlage

                    If Left(cell.Value, 3) = "BCP" And cell(1, 11).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
                    If Left(cell.Value, 3) = "BCP" And cell(1, 11).Value <= 2 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If
        
        
                      If Left(cell.Value, 3) = "BCP" And cell(1, 11).Value >= 2 Then
        cell(1, 4).Interior.ColorIndex = 46
        End If
        
Next

'--------------------------BCM---------------------------------------------
Set MyPlage = Range("A15:A" & lr)

  For Each cell In MyPlage

                    If Left(cell.Value, 3) = "BCM" And cell(1, 13).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
                    If Left(cell.Value, 3) = "BCM" And cell(1, 13).Value <= 2 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If
        
              If Left(cell.Value, 3) = "BCM" And cell(1, 13).Value >= 2 Then
        cell(1, 7).Interior.ColorIndex = 46
        End If
Next

Set MyPlage = Range("D15:D" & lr)
  For Each cell In MyPlage

                    If Left(cell.Value, 3) = "BCM" And cell(1, 11).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
                    If Left(cell.Value, 3) = "BCM" And cell(1, 11).Value <= 2 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If
        
        
                      If Left(cell.Value, 3) = "BCM" And cell(1, 11).Value >= 2 Then
        cell(1, 4).Interior.ColorIndex = 46
        End If
        
Next

'--------------------------BCG---------------------------------------------
Set MyPlage = Range("A15:A" & lr)

  For Each cell In MyPlage

                    If Left(cell.Value, 3) = "BCG" And cell(1, 13).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
                    If Left(cell.Value, 3) = "BCG" And cell(1, 13).Value <= 2 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If
        
              If Left(cell.Value, 3) = "BCG" And cell(1, 13).Value >= 2 Then
        cell(1, 7).Interior.ColorIndex = 46
        End If
Next

Set MyPlage = Range("D15:D" & lr)
  For Each cell In MyPlage

                    If Left(cell.Value, 3) = "BCG" And cell(1, 11).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
                    If Left(cell.Value, 3) = "BCG" And cell(1, 11).Value <= 2 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If
        
        
                      If Left(cell.Value, 3) = "BCG" And cell(1, 11).Value >= 2 Then
        cell(1, 4).Interior.ColorIndex = 46
        End If
        
Next

'--------------------------BCD---------------------------------------------
Set MyPlage = Range("A15:A" & lr)

  For Each cell In MyPlage

                    If Left(cell.Value, 3) = "BCD" And cell(1, 13).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
                    If Left(cell.Value, 3) = "BCD" And cell(1, 13).Value <= 2 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If
        
              If Left(cell.Value, 3) = "BCD" And cell(1, 13).Value >= 2 Then
        cell(1, 7).Interior.ColorIndex = 46
        End If
Next

Set MyPlage = Range("D15:D" & lr)
  For Each cell In MyPlage

                    If Left(cell.Value, 3) = "BCD" And cell(1, 11).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
                    If Left(cell.Value, 3) = "BCD" And cell(1, 11).Value <= 2 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If
        
        
                      If Left(cell.Value, 3) = "BCD" And cell(1, 11).Value >= 2 Then
        cell(1, 4).Interior.ColorIndex = 46
        End If
        
Next
'--------------------------BCF---------------------------------------------
Set MyPlage = Range("A15:A" & lr)

  For Each cell In MyPlage

                    If Left(cell.Value, 3) = "BCF" And cell(1, 13).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
                    If Left(cell.Value, 3) = "BCF" And cell(1, 13).Value <= 2 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If
        
              If Left(cell.Value, 3) = "BCF" And cell(1, 13).Value >= 2 Then
        cell(1, 7).Interior.ColorIndex = 46
        End If
Next

Set MyPlage = Range("D15:D" & lr)
  For Each cell In MyPlage

                    If Left(cell.Value, 3) = "BCF" And cell(1, 11).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
                    If Left(cell.Value, 3) = "BCF" And cell(1, 11).Value <= 2 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If
        
        
                      If Left(cell.Value, 3) = "BCF" And cell(1, 11).Value >= 2 Then
        cell(1, 4).Interior.ColorIndex = 46
        End If
        
Next
'--------------------------BCP---------------------------------------------
Set MyPlage = Range("A15:A" & lr)

  For Each cell In MyPlage

                    If Left(cell.Value, 3) = "BCP" And cell(1, 13).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
                    If Left(cell.Value, 3) = "BCP" And cell(1, 13).Value <= 2 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If
        
              If Left(cell.Value, 3) = "BCP" And cell(1, 13).Value >= 2 Then
        cell(1, 7).Interior.ColorIndex = 46
        End If
Next

Set MyPlage = Range("D15:D" & lr)
  For Each cell In MyPlage

                    If Left(cell.Value, 3) = "BCP" And cell(1, 11).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
                    If Left(cell.Value, 3) = "BCP" And cell(1, 11).Value <= 2 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If
        
        
                      If Left(cell.Value, 3) = "BCP" And cell(1, 11).Value >= 2 Then
        cell(1, 4).Interior.ColorIndex = 46
        End If
        
Next

'--------------------------BCZ---------------------------------------------
Set MyPlage = Range("A15:A" & lr)

  For Each cell In MyPlage

                    If Left(cell.Value, 3) = "BCZ" And cell(1, 13).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
                    If Left(cell.Value, 3) = "BCZ" And cell(1, 13).Value <= 2 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If
        
              If Left(cell.Value, 3) = "BCZ" And cell(1, 13).Value >= 2 Then
        cell(1, 7).Interior.ColorIndex = 46
        End If
Next

Set MyPlage = Range("D15:D" & lr)
  For Each cell In MyPlage

                    If Left(cell.Value, 3) = "BCZ" And cell(1, 11).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
                    If Left(cell.Value, 3) = "BCZ" And cell(1, 11).Value <= 2 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If
        
        
                      If Left(cell.Value, 3) = "BCZ" And cell(1, 11).Value >= 2 Then
        cell(1, 4).Interior.ColorIndex = 46
        End If
        
Next

'--------------------------BEF---------------------------------------------
Set MyPlage = Range("A15:A" & lr)

  For Each cell In MyPlage

                    If Left(cell.Value, 3) = "BEF" And cell(1, 13).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
                    If Left(cell.Value, 3) = "BEF" And cell(1, 13).Value <= 2 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If
        
              If Left(cell.Value, 3) = "BEF" And cell(1, 13).Value >= 2 Then
        cell(1, 7).Interior.ColorIndex = 46
        End If
Next

Set MyPlage = Range("D15:D" & lr)
  For Each cell In MyPlage

                    If Left(cell.Value, 3) = "BEF" And cell(1, 11).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
                    If Left(cell.Value, 3) = "BEF" And cell(1, 11).Value <= 2 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If
        
        
                      If Left(cell.Value, 3) = "BEF" And cell(1, 11).Value >= 2 Then
        cell(1, 4).Interior.ColorIndex = 46
        End If
        
Next

'--------------------------BER---------------------------------------------
Set MyPlage = Range("A15:A" & lr)

  For Each cell In MyPlage

                    If Left(cell.Value, 3) = "BER" And cell(1, 13).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
                    If Left(cell.Value, 3) = "BER" And cell(1, 13).Value <= 2 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If
        
              If Left(cell.Value, 3) = "BER" And cell(1, 13).Value >= 2 Then
        cell(1, 7).Interior.ColorIndex = 46
        End If
Next

Set MyPlage = Range("D15:D" & lr)
  For Each cell In MyPlage

                    If Left(cell.Value, 3) = "BER" And cell(1, 11).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
                    If Left(cell.Value, 3) = "BER" And cell(1, 11).Value <= 2 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If
        
        
                      If Left(cell.Value, 3) = "BER" And cell(1, 11).Value >= 2 Then
        cell(1, 4).Interior.ColorIndex = 46
        End If
        
Next

'--------------------------BES---------------------------------------------
Set MyPlage = Range("A15:A" & lr)

  For Each cell In MyPlage

                    If Left(cell.Value, 3) = "BES" And cell(1, 13).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
                    If Left(cell.Value, 3) = "BES" And cell(1, 13).Value <= 2 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If
        
              If Left(cell.Value, 3) = "BES" And cell(1, 13).Value >= 2 Then
        cell(1, 7).Interior.ColorIndex = 46
        End If
Next

Set MyPlage = Range("D15:D" & lr)
  For Each cell In MyPlage

                    If Left(cell.Value, 3) = "BES" And cell(1, 11).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
                    If Left(cell.Value, 3) = "BES" And cell(1, 11).Value <= 2 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If
        
        
                      If Left(cell.Value, 3) = "BES" And cell(1, 11).Value >= 2 Then
        cell(1, 4).Interior.ColorIndex = 46
        End If
        
Next

'--------------------------BAR---------------------------------------------
Set MyPlage = Range("A15:A" & lr)

  For Each cell In MyPlage

                    If Left(cell.Value, 3) = "BAR" And cell(1, 13).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
                    If Left(cell.Value, 3) = "BAR" And cell(1, 13).Value <= 2 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If
        
              If Left(cell.Value, 3) = "BAR" And cell(1, 13).Value >= 2 Then
        cell(1, 7).Interior.ColorIndex = 46
        End If
Next

Set MyPlage = Range("D15:D" & lr)
  For Each cell In MyPlage

                    If Left(cell.Value, 3) = "BAR" And cell(1, 11).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
                    If Left(cell.Value, 3) = "BAR" And cell(1, 11).Value <= 2 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If
        
        
                      If Left(cell.Value, 3) = "BAR" And cell(1, 11).Value >= 2 Then
        cell(1, 4).Interior.ColorIndex = 46
        End If
        
Next
End If

'---------------------------XDB1 ado1---------------------------------------------
'---------------------------------------------------------------------------------------------

If Error_menu.CheckBox2.Value = False And Error_menu.CheckBox1.Value = True Then

   '-------------------------Connections"----------------------------------
   
    Set MyPlage = Range("D15:D" & lr)
  For Each cell In MyPlage
  
        
        If Left(cell.Value, 4) = "XDB1" And cell(1, 11).Value > 3 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
           If Left(cell.Value, 4) = "XDB1" And cell(1, 11).Value <= 3 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If

Next
   
       Set MyPlage = Range("A15:A" & lr)
  For Each cell In MyPlage
  
        
        If Left(cell.Value, 4) = "XDB1" And cell(1, 13).Value > 3 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
        If Left(cell.Value, 4) = "XDB1" And cell(1, 13).Value <= 3 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If
        
        Next
End If

'---------------------------XDB1 connector---------------------------------------------
'---------------------------------------------------------------------------------------------
If Error_menu.CheckBox2.Value = True And Error_menu.CheckBox1.Value = False Then
Set MyPlage = Range("D15:D" & lr)
  For Each cell In MyPlage
  
        
        If Left(cell.Value, 4) = "XDB1" And cell(1, 11).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        
        End If

        
        
                If Left(cell.Value, 3) = "XDT" And cell(1, 11).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        
        End If

        
                If Left(cell.Value, 3) = "XDE" And cell(1, 11).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If


           If Left(cell.Value, 4) = "XDB1" And cell(1, 11).Value <= 2 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If
                If Left(cell.Value, 3) = "XDT" And cell(1, 11).Value <= 2 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If
                If Left(cell.Value, 3) = "XDE" And cell(1, 11).Value <= 2 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If
   
   
   Next
   
Set MyPlage = Range("A15:A" & lr)
  For Each cell In MyPlage
  
        
        If Left(cell.Value, 4) = "XDB1" And cell(1, 13).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If

                If Left(cell.Value, 3) = "XDE" And cell(1, 13).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        

                If Left(cell.Value, 3) = "XDT" And cell(1, 13).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
    
        
        If Left(cell.Value, 4) = "XDB1" And cell(1, 13).Value <= 2 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If
                If Left(cell.Value, 3) = "XDE" And cell(1, 13).Value <= 2 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If
                If Left(cell.Value, 3) = "XDT" And cell(1, 13).Value <= 2 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If
        
    Next
    End If
    
   '---------------------------------------------XDV------------------------------
Set MyPlage = Range("A15:A" & lr)
  For Each cell In MyPlage

                    If cell.Value = "XDV" And cell(1, 13).Value = 4 Then
        cell(1, 2).Interior.ColorIndex = 46
        End If

                    If cell.Value = "XDV" And cell(1, 13).Value > 4 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
                   If cell.Value = "XDV" And cell(1, 13).Value <= 3 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If

   '---------------------------------------------XDA------------------------------

                    If cell.Value = "XDÀ" And cell(1, 13).Value = 4 Then
        cell(1, 2).Interior.ColorIndex = 46
        End If

                    If cell.Value = "XDÀ" And cell(1, 13).Value > 4 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
                   If cell.Value = "XDÀ" And cell(1, 13).Value <= 3 Then
        cell(1, 2).Interior.ColorIndex = 0
        End If

Next


    
    
    

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub

