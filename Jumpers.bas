Attribute VB_Name = "Jumpers"
Sub Jumpers()
Dim lr As Long
Dim wire As Double
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False


lr = Range("A" & Rows.Count).End(xlUp).Row

'----------------- minimal wires crossection   --------------------


Set MyPlage = Range("G15:G" & lr)


wire = InputBox("Please add minimal cross-section of conductors", "Read the General Arrangement Drawings!!!", 1)


For Each cell In MyPlage
 If Not IsEmpty(cell.Value) And cell.Value < wire Then
 If Not (cell(1, 6).Value = "-" Or cell(1, 6).Value = "Shielded cable") Then
 cell.Value = wire
 cell.Font.ColorIndex = 3
 cell.Font.Bold = True
End If
End If

Next




 '------------------Jumpers Primery -------------------------
 
Set MyPlage = Range("G15:G" & lr)

    For Each cell In MyPlage

   If cell.Value = "Bridge" Then
            cell(1, 3).Value = "Insertable jumper"
            cell.Value = ""
    End If
Next





'---------------------------clear cells BAT-FCF -QAB -BGT -QCE -BCT- BCN- BAD- EB- EA- BGB-QBM2 - BPS- QBS-RLE------------------------------

Set MyPlage = Range("A15:A" & lr)

    For Each cell In MyPlage

        If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "BAT" Then
            cell(1, 9).Value = "Direct Connection"
            cell(1, 7).ClearContents
            cell(1, 8).ClearContents
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
          End If
          
                  If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "QBS" And Left(cell(1, 4).Value, 3) = "XDC" Then
           cell(1, 9).Value = "Direct Connection"
            cell(1, 7).ClearContents
            cell(1, 8).ClearContents
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
          End If
          
                  If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "BGB" Then
            cell(1, 9).Value = "Direct Connection"
            cell(1, 7).ClearContents
            cell(1, 8).ClearContents
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
          End If

                If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "QCE" Then
            cell(1, 9).Value = "Direct Connection"
            cell(1, 7).ClearContents
            cell(1, 8).ClearContents
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
          End If
          
        If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "FCF" Then
            cell(1, 9).Value = "Direct Connection"
            cell(1, 7).ClearContents
            cell(1, 8).ClearContents
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True

        End If
        
        If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "QAB" Then
            cell(1, 9).Value = "Direct Connection"
            cell(1, 7).ClearContents
            cell(1, 8).ClearContents
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
        End If
    
        If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "BGT" Then
            cell(1, 9).Value = "Direct Connection"
            cell(1, 7).ClearContents
            cell(1, 8).ClearContents
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
        End If
        
                If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "BGE" Then
            cell(1, 9).Value = "Direct Connection"
            cell(1, 7).ClearContents
            cell(1, 8).ClearContents
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
        End If
        
                If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "BCT" Then
            cell(1, 9).Value = "Direct Connection"
            cell(1, 7).ClearContents
            cell(1, 8).ClearContents
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
        End If
               If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "BCN" Then
            cell(1, 9).Value = "Direct Connection"
            cell(1, 7).ClearContents
            cell(1, 8).ClearContents
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
        End If
                 If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "BAD" Then
            cell(1, 9).Value = "Direct Connection"
            cell(1, 7).ClearContents
            cell(1, 8).ClearContents
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
        End If
        
         If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 2) = "EB" Then
          
            cell(1, 7).ClearContents
            cell(1, 8).ClearContents
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
        End If
        
         If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "EA2" Then
        
         cell(1, 7).ClearContents
        cell(1, 8).ClearContents
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
           
        If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "BPS" Then
        cell(1, 9).Value = "Direct Connection"
         cell(1, 7).ClearContents
        cell(1, 8).ClearContents
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
      
         If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "RLE" Then
         cell(1, 9).Value = "Direct Connection"
         cell(1, 7).ClearContents
        cell(1, 8).ClearContents
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
       Next
                 
 Set MyPlage = Range("D15:D1000")

    For Each cell In MyPlage
    

          If Not IsEmpty(cell(1, 4).Value) And Left(cell.Value, 3) = "BAT" Then
          cell(1, 6).Value = "Direct Connection"
            cell(1, 4).ClearContents
            cell(1, 5).ClearContents
            cell(1, 6).Font.ColorIndex = 3
            cell(1, 6).Font.Bold = True

        End If
        
        If Not IsEmpty(cell(1, 4).Value) And Left(cell.Value, 3) = "FCF" Then
        cell(1, 6).Value = "Direct Connection"
            cell(1, 4).ClearContents
            cell(1, 5).ClearContents
            cell(1, 6).Font.ColorIndex = 3
            cell(1, 6).Font.Bold = True
        End If
            If Not IsEmpty(cell(1, 4).Value) And Left(cell.Value, 3) = "QAB" Then
            cell(1, 6).Value = "Direct Connection"
            cell(1, 4).ClearContents
            cell(1, 5).ClearContents
            cell(1, 6).Font.ColorIndex = 3
            cell(1, 6).Font.Bold = True
        End If

            If Not IsEmpty(cell(1, 4).Value) And Left(cell.Value, 3) = "BGT" Then
            cell(1, 6).Value = "Direct Connection"
            cell(1, 4).ClearContents
            cell(1, 5).ClearContents
            cell(1, 6).Font.ColorIndex = 3
            cell(1, 6).Font.Bold = True
        End If
            If Not IsEmpty(cell(1, 4).Value) And Left(cell.Value, 3) = "BGE" Then
            cell(1, 6).Value = "Direct Connection"
            cell(1, 4).ClearContents
            cell(1, 5).ClearContents
            cell(1, 6).Font.ColorIndex = 3
            cell(1, 6).Font.Bold = True
        End If
            If Not IsEmpty(cell(1, 4).Value) And Left(cell.Value, 3) = "QCE" Then
            cell(1, 6).Value = "Direct Connection"
            cell(1, 4).ClearContents
            cell(1, 5).ClearContents
            cell(1, 6).Font.ColorIndex = 3
            cell(1, 6).Font.Bold = True
        End If
                    If Not IsEmpty(cell(1, 4).Value) And Left(cell.Value, 3) = "RLE" Then
            cell(1, 6).Value = "Direct Connection"
            cell(1, 4).ClearContents
            cell(1, 5).ClearContents
            cell(1, 6).Font.ColorIndex = 3
            cell(1, 6).Font.Bold = True
        End If
        
        Next
 '------------------Jumpers between equipment -------------------------
Dim x As Double
Set MyPlage = Range("A15:A" & lr)

    For Each cell In MyPlage
        
        If cell.Value <> cell(1, 4).Value Then
         If cell(1, 9).Value <> "Conductor / wire" Then
            cell(1, 9).Value = "Conductor / wire"
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
        End If
            If IsEmpty(cell(1, 7).Value) Then
             x = InputBox("Please add cross-section of conductors between" & vbNewLine & cell(1, 3) & " and " & cell(1, 6), "Wire jumper  between " & cell(1, 3) & " and " & cell(1, 6), wire, 1)
            cell(1, 7).Value = x
            cell(1, 7).Font.ColorIndex = 3
            cell(1, 7).Font.Bold = True
        End If
        If IsEmpty(cell(1, 8).Value) Then
        cell(1, 8).Value = "bk"
        cell(1, 8).Font.ColorIndex = 3
        cell(1, 8).Font.Bold = True
        End If
        End If
 Next
        
        
 '------------------XDA -------------------------

Set MyPlage = Range("A15:A" & lr)

    For Each cell In MyPlage

        

        
           If cell.Value = "XDA" And cell.Value = cell(1, 4).Value And cell(1, 9).Value = "Insertable jumper" Then
         'If Abs(cell(1, 2).Value - cell(1, 5).Value) >= 1 And cell(1, 9).Value = "Insertable jumper" Then
        cell(1, 7).ClearContents
        cell(1, 8).ClearContents
        'End If
        End If
   If cell.Value = "XDA" And cell.Value = cell(1, 4).Value And cell(1, 9).Value = "Saddle jumper" Then
         'If Abs(cell(1, 2).Value - cell(1, 5).Value) >= 1 And cell(1, 9).Value = "Saddle jumper" Then
        cell(1, 7).ClearContents
        cell(1, 8).ClearContents
        cell(1, 9).Value = "Insertable jumper"
        cell(1, 9).Font.ColorIndex = 6
        cell(1, 9).Font.Bold = True
        'End If
        End If
        
        


        
        
'-------------Black jumpers----------------------------------
        If cell.Value = "XDA" And cell.Value = cell(1, 4).Value And Not (cell(1, 9).Value = "Insertable jumper" Or cell(1, 9).Value = "Saddle jumper") Then
        cell(1, 7).ClearContents
        cell(1, 8).ClearContents
        cell(1, 9).Value = "Insertable jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
     
    Next
        

        
     '------------------XDC -------------------------

    
Set MyPlage = Range("A15:A" & lr)

    For Each cell In MyPlage


        

        
       If Left(cell.Value, 3) = "XDC" And (cell(1, 9).Value = "Wire jumper" Or cell(1, 9).Value = "Conductor / wire") Then
       If IsEmpty(cell(1, 7).Value) Then
        XDC = InputBox("Please add cross-section of conductors between" & vbNewLine & cell(1, 3) & " and " & cell(1, 6), "Wire connection" & cell(1, 3) & " and " & cell(1, 6), wire)
        cell(1, 7).Value = XDC
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        End If
        If IsEmpty(cell(1, 8).Value) Then
        cell(1, 8).Value = "bk"
        cell(1, 8).Font.ColorIndex = 3
        cell(1, 8).Font.Bold = True
        End If
        End If
        
    If Left(cell.Value, 3) = "XDC" And cell.Value = cell(1, 4).Value Then
    If Abs(cell(1, 2).Value - cell(1, 5).Value) >= 1 And (cell(1, 9).Value = "Saddle jumper" Or cell(1, 9).Value = "Insertable jumper") Then
    answer = MsgBox("Is connection between " & cell(1, 3) & " and " & cell(1, 6) & " is with -" & cell(1, 9), vbYesNo + vbQuestion + vbDefaultButton2, "XDC matal jumper")
    If answer = vbNo Then
    XDC = InputBox("Please add cross-section of conductors between" & cell(1, 3) & " and " & cell(1, 6), "Wire jumper  between " & cell(1, 3) & " and " & cell(1, 6), wire)
        If XDC = vbOK Then
        If IsEmpty(cell(1, 7).Value) Then
        cell(1, 7).Value = XDC
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        End If
        If IsEmpty(cell(1, 8).Value) Then
        cell(1, 8).Value = "bk"
        cell(1, 8).Font.ColorIndex = 3
        cell(1, 8).Font.Bold = True
        End If
        If Not cell(1, 9).Value = "Wire jumper" Then
        cell(1, 9).Value = "Wire jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
        Else
        cell(1, 7).ClearContents
        cell(1, 8).ClearContents
        If Not cell(1, 9).Value = "Saddle jumper" Then
        cell(1, 9).Value = "Saddle jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
        End If
        End If
        


        If IsEmpty(cell(1, 8).Value) And Left(cell.Value, 3) = "XDC" And cell.Value = cell(1, 4).Value And Not (cell(1, 9).Value = "Saddle jumper" Or cell(1, 9).Value = "Insertable jumper") Then
        If Abs(cell(1, 2).Value - cell(1, 5).Value) > 1 Then
        cell(1, 8).Value = "bk"
        cell(1, 8).Font.ColorIndex = 3
        cell(1, 8).Font.Bold = True
        End If
        End If

    Next

 
     '------------------XDM -------------------------

Set MyPlage = Range("A15:A" & lr)

    For Each cell In MyPlage

        

        
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

Set MyPlage = Range("A15:A" & lr)

    For Each cell In MyPlage

        
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

Set MyPlage = Range("A15:A" & lr)

    For Each cell In MyPlage

        
        
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

Set MyPlage = Range("A15:A" & lr)

    For Each cell In MyPlage

        
        
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

Set MyPlage = Range("A15:A" & lr)

    For Each cell In MyPlage

        

        
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

Set MyPlage = Range("A15:A" & lr)

    For Each cell In MyPlage


            If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "XDI" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) = 1 And cell(1, 9).Value = "Conductor / wire" Then
        answer = MsgBox("Is connection between" & cell(1, 3).Value & " And " & cell(1, 6).Value & " is " & cell(1, 9).Value, vbYesNo + vbQuestion, "XDI jumper")
        If answer = vbNo Then
        cell(1, 7).ClearContents
        cell(1, 8).ClearContents
        cell(1, 9).Value = "Saddle jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        Else
        End If
        End If
                  If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "XDI" And cell.Value = cell(1, 4).Value Then
                  If Abs(cell(1, 2).Value - cell(1, 5).Value) = 1 And cell(1, 9).Value = "Wire jumper" Then
                  answer = MsgBox("Is connection between" & cell(1, 3).Value & " And " & cell(1, 6).Value & " is " & cell(1, 9).Value, vbYesNo + vbQuestion, "XDI jumper")
        If answer = vbNo Then
        cell(1, 7).ClearContents
        cell(1, 8).ClearContents
        cell(1, 9).Value = "Saddle jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        Else
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
        
           If IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "XDI" And (cell(1, 9).Value = "Wire jumper" Or cell(1, 9).Value = "Ponticello a filo") Then
        XDI = InputBox("Please add cross-section of conductors between" & cell(1, 3) & " and " & cell(1, 6), "Wire jumper between " & cell(1, 3) & " and " & cell(1, 6), wire)
        cell(1, 7).Value = XDI
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        'cell(1, 8).Value = "bk"
        End If
        
        
           If IsEmpty(cell(1, 8).Value) And Left(cell.Value, 3) = "XDI" And cell.Value = cell(1, 4).Value Then
        If Abs(cell(1, 2).Value - cell(1, 5).Value) > 1 Then
        cell(1, 8).Value = "bk"
        cell(1, 8).Font.ColorIndex = 3
        cell(1, 8).Font.Bold = True
        End If
        End If

        
        
        Next
    '--------------------------------------RAR------------------------------
    Set MyPlage = Range("A15:A" & lr)

    For Each cell In MyPlage

          If Left(cell.Value, 3) = "RAR" Then
          answer = MsgBox("Do you want to clear wire connection between " & vbNewLine & cell(1, 3).Value & " and " & cell(1, 6).Value, vbYesNo + vbQuestion, "RAR section")
          If answer = vbYes Then
            cell(1, 7).ClearContents
            cell(1, 8).ClearContents
         End If
            End If
 Next
                 
                 
 Set MyPlage = Range("D15:D1000")

    For Each cell In MyPlage
    

          If Left(cell.Value, 3) = "RAR" Then
          answer = MsgBox("Do you want to clear wire connection between " & vbNewLine & cell(1, 3).Value & " and " & cell.Offset(1, -1).Value, vbYesNo + vbQuestion, "RAR section")
        If answer = vbYes Then
            cell(1, 4).ClearContents
            cell(1, 5).ClearContents
        End If
        End If
Next

'------------------XDX -------------------------

Set MyPlage = Range("A15:A" & lr)

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
        
        
                If IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "XDX" And (cell(1, 9).Value = "Wire jumper" Or cell(1, 9).Value = "Ponticello a filo") Then
        XDX = InputBox("Please add cross-section of conductors between" & vbNewLine & cell(1, 3) & " and " & cell(1, 6), "Wire jumper between " & cell(1, 3) & " and " & cell(1, 6), wire)
        cell(1, 7).Value = XDX
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        'cell(1, 8).Value = "bk"
        End If
        

           If IsEmpty(cell(1, 8).Value) And Left(cell.Value, 3) = "XDX" And cell.Value = cell(1, 4).Value Then
        If Abs(cell(1, 2).Value - cell(1, 5).Value) > 1 Then
        cell(1, 8).Value = "bk"
        cell(1, 8).Font.ColorIndex = 3
        cell(1, 8).Font.Bold = True
        End If
        End If
        
 
 
  '---------------------------Wire Bridges----------------------------------------------
  
  


        
           If cell.Value = "XDB1" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) >= 1 And cell(1, 9).Value = "Insertable jumper" Then
        cell(1, 9).Value = "Wire jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        If IsEmpty(cell(1, 8).Value) Then
        cell(1, 8).Value = "bk"
        cell(1, 8).Font.Bold = True
        cell(1, 8).Font.ColorIndex = 3
        End If
        End If
        End If
   If cell.Value = "XDB1" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) >= 1 And cell(1, 9).Value = "Saddle jumper" Then
        cell(1, 9).Value = "Wire jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        If IsEmpty(cell(1, 8).Value) Then
        cell(1, 8).Value = "bk"
        cell(1, 8).Font.ColorIndex = 3
        cell(1, 8).Font.Bold = True
        End If
        End If
        End If
        
          If IsEmpty(cell(1, 7).Value) And cell.Value = "XDB1" And (cell(1, 9).Value = "Wire jumper" Or cell(1, 9).Value = "Ponticello a filo") Then
        XDB1 = InputBox("Please add cross-section of conductors between" & vbNewLine & cell(1, 3) & " and " & cell(1, 6), "Wire jumper between " & cell(1, 3) & " and " & cell(1, 6), wire)
        cell(1, 7).Value = XDB1
        'cell(1, 8).Value = "bk"
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        End If
        
            If IsEmpty(cell(1, 8).Value) And cell.Value = "XDB1" And cell.Value = cell(1, 4).Value Then
        If Abs(cell(1, 2).Value - cell(1, 5).Value) >= 1 Then
        cell(1, 8).Value = "bk"
        cell(1, 8).Font.ColorIndex = 3
        cell(1, 8).Font.Bold = True
        End If
        End If
        
        

        
           If cell.Value = "XDE" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) >= 1 And cell(1, 9).Value = "Insertable jumper" Then
        cell(1, 9).Value = "Wire jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
   If cell.Value = "XDE" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) >= 1 And cell(1, 9).Value = "Saddle jumper" Then
        cell(1, 9).Value = "Wire jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
        
        
                    If IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "XDE" And (cell(1, 9).Value = "Wire jumper" Or cell(1, 9).Value = "Ponticello a filo") Then
        XDE = InputBox("Please add cross-section of conductors between" & vbNewLine & cell(1, 3) & " and " & cell(1, 6), "Wire jumper between " & cell(1, 3) & " and " & cell(1, 6), wire)
        cell(1, 7).Value = XDE
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        
        'cell(1, 8).Value = "bk"
        End If
              
        
        
        
           If cell.Value = "XDT" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) >= 1 And cell(1, 9).Value = "Insertable jumper" Then
        cell(1, 9).Value = "Wire jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
   If cell.Value = "XDT" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) >= 1 And cell(1, 9).Value = "Saddle jumper" Then
        cell(1, 9).Value = "Wire jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
                  If IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "XDT" And (cell(1, 9).Value = "Wire jumper" Or cell(1, 9).Value = "Ponticello a filo") Then
        XDT = InputBox("Please add cross-section of conductors between" & vbNewLine & cell(1, 3) & " and " & cell(1, 6), "Wire jumper between " & cell(1, 3) & " and " & cell(1, 6), wire)
        cell(1, 7).Value = XDT
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        'cell(1, 8).Value = "bk"
        End If
        
                   If IsEmpty(cell(1, 8).Value) And (Left(cell.Value, 3) = "XDT" Or Left(cell.Value, 3) = "XDE") And cell.Value = cell(1, 4).Value Then
        If Abs(cell(1, 2).Value - cell(1, 5).Value) >= 1 Then
        cell(1, 8).Value = "bk"
        cell(1, 8).Font.ColorIndex = 3
        cell(1, 8).Font.Bold = True
        End If
        End If
        
        
        
  Next
 
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True


        
 End Sub
  
   


   
