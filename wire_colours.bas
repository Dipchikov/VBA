Attribute VB_Name = "wire_colours"
Sub wire_colours()
Dim lr As Long

lr = Range("A" & Rows.Count).End(xlUp).Row
'Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
    '-----------------Colours of cable---Interconection-----------------------------
   Set MyPlage = Range("G12:G" & lr)

    For Each cell In MyPlage
    
        If cell.Value = "bn" Or cell.Value = "BN" Then
            cell(1, 4).Interior.ColorIndex = 53 ' old code 46
        End If
        
           If cell.Value = "bu" Or cell.Value = "BU" Then
            cell(1, 4).Interior.ColorIndex = 32
        End If
        
              If cell.Value = "lbu" Or cell.Value = "LBU" Then
            cell(1, 4).Interior.ColorIndex = 33

            
        End If
        
        If cell.Value = "gr" Or cell.Value = "GR" Then
            cell(1, 4).Interior.ColorIndex = 15

        End If
        
        If cell.Value = "gy" Or cell.Value = "GY" Then
            cell(1, 4).Interior.ColorIndex = 15

        End If

        If cell.Value = "rd" Or cell.Value = "RD" Then
            cell(1, 4).Interior.ColorIndex = 3

        End If
        
             If cell.Value = "vt" Or cell.Value = "VT" Then
            cell(1, 4).Interior.ColorIndex = 39

        End If
        
                     If cell.Value = "og" Or cell.Value = "OG" Then
            cell(1, 4).Interior.ColorIndex = 44

        End If
           
    Next
     '-----------------Lolours of cable---WCT-----------------------------
     
      '-----------------Shielded cable--------------------------------
    
    Set MyPlage = Range("L15:L" & lr)

    For Each cell In MyPlage
    
        If cell.Value = "Shielded cable" Then
            cell.Interior.ColorIndex = 6
            cell(1, -1).Interior.ColorIndex = 6
            cell(1, -2).Interior.ColorIndex = 6
            cell(1, -3).Interior.ColorIndex = 6
            cell(1, -4).Interior.ColorIndex = 6
        End If

    Next
     
     
   Set MyPlage = Range("H12:H" & lr)

    For Each cell In MyPlage
    
        If cell.Value = "bn" Or cell.Value = "BN" Then
            cell(1, 5).Interior.ColorIndex = 53
        End If
        
           If cell.Value = "bu" Or cell.Value = "BU" Then
            cell(1, 5).Interior.ColorIndex = 32

            
        End If
        
        
              If cell.Value = "lbu" Or cell.Value = "LBU" Then
            cell(1, 5).Interior.ColorIndex = 33

            
        End If
        
        If cell.Value = "gr" Or cell.Value = "GR" Then
            cell(1, 5).Interior.ColorIndex = 15

        End If
        
        If cell.Value = "gy" Or cell.Value = "GY" Then
            cell(1, 5).Interior.ColorIndex = 15

        End If

        If cell.Value = "rd" Or cell.Value = "RD" Then
            cell(1, 5).Interior.ColorIndex = 3

        End If
        
             If cell.Value = "vt" Or cell.Value = "VT" Then
            cell(1, 5).Interior.ColorIndex = 39

        End If
        
                     If cell.Value = "og" Or cell.Value = "OG" Then
            cell(1, 5).Interior.ColorIndex = 44

        End If
           
    Next
      'Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    End Sub
