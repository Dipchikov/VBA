Attribute VB_Name = "ErrorsRefs"
Sub ErrorsRefs()
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False

Set MyPlage = Range("A15:A1000")

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

Set MyPlage = Range("D15:D1000")
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
Set MyPlage = Range("A15:A1000")

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

Set MyPlage = Range("D15:D1000")
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
Set MyPlage = Range("A15:A1000")

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

Set MyPlage = Range("D15:D1000")
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
Set MyPlage = Range("A15:A1000")

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

Set MyPlage = Range("D15:D1000")
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
Set MyPlage = Range("A15:A1000")

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

Set MyPlage = Range("D15:D1000")
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
Set MyPlage = Range("A15:A1000")

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

Set MyPlage = Range("D15:D1000")
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
Set MyPlage = Range("A15:A1000")

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

Set MyPlage = Range("D15:D1000")
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
Set MyPlage = Range("A15:A1000")

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

Set MyPlage = Range("D15:D1000")
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
Set MyPlage = Range("A15:A1000")

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

Set MyPlage = Range("D15:D1000")
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
Set MyPlage = Range("A15:A1000")

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

Set MyPlage = Range("D15:D1000")
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
Set MyPlage = Range("A15:A1000")

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

Set MyPlage = Range("D15:D1000")
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
Set MyPlage = Range("A15:A1000")

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

Set MyPlage = Range("D15:D1000")
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
Set MyPlage = Range("A15:A1000")

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

Set MyPlage = Range("D15:D1000")
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
Set MyPlage = Range("A15:A1000")

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

Set MyPlage = Range("D15:D1000")
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

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

