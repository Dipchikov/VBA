Attribute VB_Name = "Error_number_of_conections"
Sub Error_number_of_conections()

Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
'---------------------------------------------XDC------------------------------
Set MyPlage = Range("A15:A1000")
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
Set MyPlage = Range("D15:D1000")
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
Set MyPlage = Range("A15:A1000")
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


Set MyPlage = Range("D15:D1000")
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
Set MyPlage = Range("A15:A1000")
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
Set MyPlage = Range("D15:D1000")
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
Set MyPlage = Range("A15:A1000")
  For Each cell In MyPlage

        
                   If Left(cell.Value, 3) = "FCM" And cell(1, 13).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
                          
Next

Set MyPlage = Range("D15:D1000")
  For Each cell In MyPlage

        
           If Left(cell.Value, 3) = "FCM" And cell(1, 11).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
        
Next
      
'---------------------------------------------KFA to KFZ------------------------------
Set MyPlage = Range("A15:A1000")
  For Each cell In MyPlage

        
                   If Left(cell.Value, 2) = "KF" And cell(1, 13).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If

Next

Set MyPlage = Range("D15:D1000")
  For Each cell In MyPlage

        
           If Left(cell.Value, 2) = "KF" And cell(1, 11).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        

  Next

'---------------------------------------------Lamps------------------------------
Set MyPlage = Range("A15:A1000")
  For Each cell In MyPlage

        
                   If Left(cell.Value, 3) = "PFB" And cell(1, 13).Value > 4 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
                
                   If Left(cell.Value, 3) = "PFG" And cell(1, 13).Value > 4 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
                       If Left(cell.Value, 3) = "PFR" And cell(1, 13).Value > 4 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
             If Left(cell.Value, 3) = "PFY" And cell(1, 13).Value > 4 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
        If Left(cell.Value, 3) = "SPM" And cell(1, 13).Value > 4 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
        If Left(cell.Value, 3) = "SFT" And cell(1, 13).Value > 4 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
        If Left(cell.Value, 3) = "STF" And cell(1, 13).Value > 4 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If

        If Left(cell.Value, 3) = "PFL" And cell(1, 13).Value > 4 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
Next

Set MyPlage = Range("D15:D1000")
  For Each cell In MyPlage

        
           If Left(cell.Value, 3) = "PFB" And cell(1, 11).Value > 4 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
           If Left(cell.Value, 3) = "PFG" And cell(1, 11).Value > 4 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
              If Left(cell.Value, 3) = "PFR" And cell(1, 11).Value > 4 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
                      If Left(cell.Value, 3) = "PFY" And cell(1, 11).Value > 4 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
            
     
                If Left(cell.Value, 3) = "SPM" And cell(1, 11).Value > 4 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
              If Left(cell.Value, 3) = "SFT" And cell(1, 11).Value > 4 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
             If Left(cell.Value, 3) = "STF" And cell(1, 11).Value > 4 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
             
             If Left(cell.Value, 3) = "PFL" And cell(1, 11).Value > 4 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
     
Next

'--------------------------------------Selector Switch------------------------------
Set MyPlage = Range("A15:A1000")
  For Each cell In MyPlage

        If Left(cell.Value, 2) = "SF" And cell(1, 13).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
Next

Set MyPlage = Range("D15:D1000")
  For Each cell In MyPlage

        
         If Left(cell.Value, 2) = "SF" And cell(1, 11).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
Next

'--------------------------------------BT- Thermostat------------------------------
Set MyPlage = Range("A15:A1000")
  For Each cell In MyPlage

        If Left(cell.Value, 2) = "BT" And cell(1, 13).Value > 1 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
Next

Set MyPlage = Range("D15:D1000")
  For Each cell In MyPlage

        
         If Left(cell.Value, 2) = "BT" And cell(1, 11).Value > 1 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
Next
'--------------------------------------XDM------------------------------
Set MyPlage = Range("A15:A1000")
  For Each cell In MyPlage

        If Left(cell.Value, 3) = "XDM" And cell(1, 13).Value > 1 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
Next

Set MyPlage = Range("D15:D1000")
  For Each cell In MyPlage

        
         If Left(cell.Value, 3) = "XDM" And cell(1, 11).Value > 1 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
Next
'--------------------------------------PFV------------------------------
Set MyPlage = Range("A15:A1000")
  For Each cell In MyPlage

        If Left(cell.Value, 3) = "PFV" And cell(1, 13).Value > 1 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
Next

Set MyPlage = Range("D15:D1000")
  For Each cell In MyPlage

        
         If Left(cell.Value, 3) = "PFV" And cell(1, 11).Value > 1 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
    Next
'--------------------------------------PGA to PGW------------------------------
Set MyPlage = Range("A15:A1000")
  For Each cell In MyPlage

        If Left(cell.Value, 2) = "PG" And cell(1, 13).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
Next

Set MyPlage = Range("D15:D1000")
  For Each cell In MyPlage

        
         If Left(cell.Value, 2) = "PG" And cell(1, 11).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
Next

'---------------------------------------------PGM------------------------------
Set MyPlage = Range("A15:A1000")
  For Each cell In MyPlage

        
                   If Left(cell.Value, 3) = "PGM" And cell(1, 13).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
                          
Next

Set MyPlage = Range("D15:D1000")
  For Each cell In MyPlage

        
           If Left(cell.Value, 3) = "PGM" And cell(1, 11).Value > 2 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
        
        
Next
             
 '--------------------------------------RAR------------------------------
Set MyPlage = Range("A15:A1000")
  For Each cell In MyPlage

        If Left(cell.Value, 3) = "RAR" And cell(1, 13).Value > 3 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
Next

Set MyPlage = Range("D15:D1000")
  For Each cell In MyPlage

        
         If Left(cell.Value, 2) = "RAR" And cell(1, 11).Value > 3 Then
        cell(1, 2).Interior.ColorIndex = 3
        End If
Next
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub

