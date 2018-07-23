Attribute VB_Name = "ERRORS"
Public Sub Errors()
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False

Dim XDA As Single
Dim XDV As Single
'----------default values----------
XDA = XDA1
XDV = XDV1
'------------------XDA -------------------------
'Range("G7:H1000").Interior.ColorIndex = 0
Set MyPlage = Range("A15:A1000")
  For Each cell In MyPlage
  
        
    If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "XDA" And cell(1, 7).Value <> XDA Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = XDA
        End If
        
    Next
    Set MyPlage = Range("D15:D1000")
  For Each cell In MyPlage
  
        
            If Not IsEmpty(cell(1, 4).Value) And Left(cell.Value, 3) = "XDA" And cell(1, 4).Value <> XDA Then
        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = XDA
        End If


 Next
         
 
 '------------------XDV -------------------------

Set MyPlage = Range("A15:A1000")
  For Each cell In MyPlage
  
        
            If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "XDV" And cell(1, 7).Value <> XDV Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = XDV
      
        End If
              
        Next
    Set MyPlage = Range("D15:D1000")
  For Each cell In MyPlage
  
        
            If Not IsEmpty(cell(1, 4).Value) And Left(cell.Value, 3) = "XDV" And cell(1, 4).Value <> XDV Then
        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = XDV
         End If
                     
         Next
         
     
        '---------------------------FCM----------------------------------------------
        Set MyPlage = Range("A15:A1000")
        For Each cell In MyPlage
        
                If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 3 And cell(1, 4).Value = "XDI1" And cell(1, 7).Value < 2.5 Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = 2.5
        End If
        
        If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 3 And cell(1, 4).Value = "XDI2" And cell(1, 7).Value < 2.5 Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = motor
        End If

        
        If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 3 And cell(1, 4).Value = "XDI3" And cell(1, 7).Value < 2.5 Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = 2.5
        End If
       
        If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 3 And cell(1, 4).Value = "XDI4" And cell(1, 7).Value < 2.5 Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = 2.5
        End If
       
        If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 3 And cell(1, 4).Value = "XDI5" And cell(1, 7).Value < 2.5 Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = 2.5
        End If
        
        If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 3 And cell(1, 4).Value = "XDI7" And cell(1, 7).Value < 2.5 Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = 2.5
        End If
       
        If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 3 And cell(1, 4).Value = "XDI8" And cell(1, 7).Value < 2.5 Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = 2.5
        End If
        
        If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 3 And cell(1, 4).Value = "XDI9" And cell(1, 7).Value < 2.5 Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = 2.5
        End If
        
        
        
         If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 1 And cell(1, 4).Value = "XDI1" And cell(1, 7).Value < 2.5 Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = 2.5
        End If
        
        If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 1 And cell(1, 4).Value = "XDI2" And cell(1, 7).Value < 2.5 Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = 2.5
        End If

        
        If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 1 And cell(1, 4).Value = "XDI3" And cell(1, 7).Value < 2.5 Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = 2.5
        End If
       
        If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 1 And cell(1, 4).Value = "XDI4" And cell(1, 7).Value < 2.5 Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = 2.5
        End If
       
        If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 1 And cell(1, 4).Value = "XDI5" And cell(1, 7).Value < 2.5 Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = 2.5
        End If
        
        If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 1 And cell(1, 4).Value = "XDI7" And cell(1, 7).Value < 2.5 Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = 2.5
        End If
       
        If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 1 And cell(1, 4).Value = "XDI8" And cell(1, 7).Value < 2.5 Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = 2.5
        End If
        
        If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 1 And cell(1, 4).Value = "XDI9" And cell(1, 7).Value < 2.5 Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = 2.5
        End If
        

Next
 
'---------------------------XDI6----------------------------------------------
        
         Set MyPlage = Range("A15:A1000")
        For Each cell In MyPlage
        
        If Not IsEmpty(cell(1, 7).Value) And cell.Value = "XDI6" And cell(1, 7).Value <> XDV Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = XDV
        End If
        Next
 
    'If Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 3 And cell(1, 4).Value = "XDI6" And cell(1, 7).Value <> XDV Then
       ' cell(1, 7).Font.ColorIndex = 3
       '' cell(1, 7).Font.Bold = True
       ' cell(1, 7).Value = XDV
       ' End If
        'Next
        
    
 Set MyPlage = Range("D15:d1000")
  For Each cell In MyPlage
  
        
If Not IsEmpty(cell(1, 4).Value) And cell.Value = "XDI6" And cell(1, 4).Value <> XDV Then
        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = XDV
        
 End If
                
Next
         
 '---------------------------XDI8----------------------------------------------
    Set MyPlage = Range("A15:A1000")
  For Each cell In MyPlage
  
            If Not IsEmpty(cell(1, 7).Value) And cell.Value = "XDI8" And cell(1, 7).Value <> XDA Then
            If Not Left(cell(1, 5).Value, 1) = "A" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = XDA
        End If
        End If
        
            If Not IsEmpty(cell(1, 7).Value) And cell(1, 4).Value = "XDI8" And cell(1, 7).Value <> XDA Then
            If Not Left(cell(1, 2).Value, 1) = "A" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = XDA
        End If
        End If
        
        
        
 Next
    
  
   '---------------------------XDI1----------------------------------------------
  'Set MyPlage = Range("A15:A1000")
  'For Each cell In MyPlage
  
        
           ' If Not IsEmpty(cell(1, 7).Value) And cell.Value = "XDI1" And cell(1, 7).Value < 2.5 Then
            'If Not Left(cell(1, 5).Value, 1) = "A" Then
       ' cell(1, 7).Font.ColorIndex = 3
       ' cell(1, 7).Font.Bold = True
       ' cell(1, 7).Value = 2.5
       ' End If
       ' End If
        
            ' If Not IsEmpty(cell(1, 7).Value) And cell.Value = "XDI1" And cell(1, 7).Value < 2.5 And cell(1, 13).Value >= 2 Then
        'cell(1, 7).Font.ColorIndex = 3
       ' cell(1, 7).Font.Bold = True
       '' cell(1, 7).Value = 2.5
       ' End If
        

        
             'If Not IsEmpty(cell(1, 7).Value) And cell(1, 4).Value = "XDI1" And cell(1, 7).Value < 2.5 Then
            'If Not Left(cell(1, 2).Value, 1) = "A" Then
        'cell(1, 7).Font.ColorIndex = 3
       ' cell(1, 7).Font.Bold = True
       ' cell(1, 7).Value = 2.5
       ' End If
       ' End If
        
        
            ' If Not IsEmpty(cell(1, 7).Value) And cell(1, 4).Value = "XDI1" And cell(1, 7).Value < 2.5 And cell(1, 14).Value >= 2 Then
       ' cell(1, 7).Font.ColorIndex = 3
        'cell(1, 7).Font.Bold = True
        'cell(1, 7).Value = 2.5
       ' End If
        
 
   ' Next
  
    

   '---------------------------XDI2----------------------------------------------
  Set MyPlage = Range("A15:A1000")
  For Each cell In MyPlage
  
        
            If Not IsEmpty(cell(1, 7).Value) And cell.Value = "XDI2" And cell(1, 7).Value < 2.5 Then
            If Not Left(cell(1, 5).Value, 1) = "A" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = 2.5
        End If
        End If
        
             'If Not IsEmpty(cell(1, 7).Value) And cell.Value = "XDI2" And cell(1, 7).Value < 2.5 And cell(1, 13).Value >= 2 Then
        'cell(1, 7).Font.ColorIndex = 3
        'cell(1, 7).Font.Bold = True
        'cell(1, 7).Value = 2.5
       ' End If
        '

        
             If Not IsEmpty(cell(1, 7).Value) And cell(1, 4).Value = "XDI2" And cell(1, 7).Value < 2.5 Then
            If Not Left(cell(1, 2).Value, 1) = "A" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = 2.5
        End If
        End If
        
        ' If Not IsEmpty(cell(1, 7).Value) And cell(1, 4).Value = "XDI2" And cell(1, 7).Value < 2.5 And cell(1, 14).Value >= 2 Then
       ' cell(1, 7).Font.ColorIndex = 3
       ' cell(1, 7).Font.Bold = True
       ' cell(1, 7).Value = 2.5
       ' End If
 
    Next


         
         
      '---------------------------XDI3----------------------------------------------
Set MyPlage = Range("A15:A1000")
  For Each cell In MyPlage
  
        
            If Not IsEmpty(cell(1, 7).Value) And cell.Value = "XDI3" And cell(1, 7).Value < 2.5 Then
            If Not Left(cell(1, 5).Value, 1) = "A" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = 2.5
        End If
        End If
        
         ' If Not IsEmpty(cell(1, 7).Value) And cell.Value = "XDI3" And cell(1, 7).Value < 2.5 And cell(1, 13).Value >= 2 Then
       ' cell(1, 7).Font.ColorIndex = 3
       ' cell(1, 7).Font.Bold = True
       ' cell(1, 7).Value = 2.5
      '  End If
        
             If Not IsEmpty(cell(1, 7).Value) And cell(1, 4).Value = "XDI3" And cell(1, 7).Value < 2.5 Then
            If Not Left(cell(1, 2).Value, 1) = "A" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = 2.5
        End If
        End If
    'If Not IsEmpty(cell(1, 7).Value) And cell(1, 4).Value = "XDI3" And cell(1, 7).Value < 2.5 And cell(1, 14).Value >= 2 Then
        'cell(1, 7).Font.ColorIndex = 3
        'cell(1, 7).Font.Bold = True
        'cell(1, 7).Value = 2.5
        'End If

Next
         
      '--------------------------------PGA--------------------------------------------


Set MyPlage = Range("A15:A1000")
  For Each cell In MyPlage
  
        
            If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "PGA" And cell(1, 7).Value <> XDA Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = XDA
        End If
                
 Next
 
    Set MyPlage = Range("D15:D1000")
  For Each cell In MyPlage
  
        
            If Not IsEmpty(cell(1, 4).Value) And Left(cell.Value, 3) = "PGA" And cell(1, 4).Value <> XDA Then
        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = XDA
        
        End If
               
Next
   
   
   '--------------------------------PGV--------------------------------------------


Set MyPlage = Range("A15:A1000")
  For Each cell In MyPlage
  
        
            If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "PGV" And cell(1, 7).Value <> XDV Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = XDV
        End If
Next
        
    Set MyPlage = Range("D15:D1000")
  For Each cell In MyPlage
  
        
            If Not IsEmpty(cell(1, 4).Value) And Left(cell.Value, 3) = "PGV" And cell(1, 4).Value <> XDV Then

        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = XDV
        
        End If

 Next
 

'--------------------------------XDMs-Eroors-------------------------------------------

XDMs_errors.XDMs_errors


'-------------------------XE"----------------------------------

 Set MyPlage = Range("A15:A1000")
  For Each cell In MyPlage
  
        
        If Not IsEmpty(cell(1, 7).Value) And Not (cell(1, 8).Value = "gnye" Or cell(1, 8).Value = "GNYE") And Left(cell.Value, 2) = "XE" Then
        cell(1, 8).Value = "gnye"
        cell(1, 8).Font.ColorIndex = 3
        cell(1, 8).Font.Bold = True
        End If

        If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 2) = "XE" And cell(1, 7).Value < 2.5 Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = 2.5
        End If
        
                
        If Not IsEmpty(cell(1, 7).Value) And Not (cell(1, 8).Value = "gnye" Or cell(1, 8).Value = "GNYE") And Left(cell.Value, 2) = "PE" Then
        cell(1, 8).Value = "gnye"
        cell(1, 8).Font.ColorIndex = 3
        cell(1, 8).Font.Bold = True
        End If

        If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 2) = "PE" And cell(1, 7).Value < 2.5 Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = 2.5
        End If
                 
        If Not IsEmpty(cell(1, 7).Value) And Not (cell(1, 8).Value = "gnye" Or cell(1, 8).Value = "GNYE") And Left(cell.Value, 2) = "IE" Then
        cell(1, 8).Value = "gnye"
        cell(1, 8).Font.ColorIndex = 3
        cell(1, 8).Font.Bold = True
        End If

        If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 2) = "IE" And cell(1, 7).Value < 2.5 Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = 2.5
        End If
        Next


   
    Set MyPlage = Range("D15:D1000")
  For Each cell In MyPlage
  
        
        If Not IsEmpty(cell(1, 4).Value) And Not (cell(1, 5).Value = "gnye" Or cell(1, 5).Value = "GNYE") And Left(cell.Value, 2) = "XE" Then
        cell(1, 5).Value = "gnye"
        cell(1, 5).Font.ColorIndex = 3
        cell(1, 5).Font.Bold = True
        End If

        If Not IsEmpty(cell(1, 4).Value) And Left(cell.Value, 2) = "XE" And cell(1, 4).Value < 2.5 Then
        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = 2.5
        End If
        
                
        If Not IsEmpty(cell(1, 4).Value) And Not (cell(1, 5).Value = "gnye" Or cell(1, 5).Value = "GNYE") And Left(cell.Value, 2) = "PE" Then
        cell(1, 5).Value = "gnye"
        cell(1, 5).Font.ColorIndex = 3
        cell(1, 5).Font.Bold = True
        End If

        If Not IsEmpty(cell(1, 4).Value) And Left(cell.Value, 2) = "PE" And cell(1, 4).Value < 2.5 Then
        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = 2.5
        End If
                 
        If Not IsEmpty(cell(1, 4).Value) And Not (cell(1, 5).Value = "gnye" Or cell(1, 5).Value = "GNYE") And Left(cell.Value, 2) = "IE" Then
        cell(1, 5).Value = "gnye"
        cell(1, 5).Font.ColorIndex = 3
        cell(1, 5).Font.Bold = True
        End If

        If Not IsEmpty(cell(1, 4).Value) And Left(cell.Value, 2) = "IE" And cell(1, 4).Value < 2.5 Then
        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = 2.5
        End If
Next
   

   
   '------------------XDB93 -XDB91----------------------------------
Set MyPlage = Range("A15:A1000")
  For Each cell In MyPlage
  
        
If cell.Value = "XDB91" Then
answer = MsgBox("Please check for other connection between XDB93 and XDB91!!!", vbOKOnly + vbExclamation, "Connection between XDB93 and XDB91")
 End If

          
Next
Set MyPlage = Range("D15:D1000")
  For Each cell In MyPlage
  
 If cell.Value = "XDB91" Then
answer = MsgBox("Please check for other connection between XDB93 and XDB91!!!", vbOKOnly + vbExclamation, "Connection between XDB93 and XDB91")
 End If

          
Next



 '------------------XDV -------------------------
Set MyPlage = Range("A15:A1000")
  For Each cell In MyPlage
  
        
            If cell.Value = "XDV" And cell.Value = cell(1, 4).Value And cell(1, 2).Value = 10 And cell(1, 5).Value = 11 And Not (cell(1, 9).Value = "Wire jumper" Or cell(1, 9).Value = "Ponticello a filo") Then
            answer = MsgBox("Is the connection between XDV:10 and XDV:11  is with wire jumper ?", vbYesNo + vbQuestion, "-XDV jumpers")
            If answer = vbYes Then
            cell(1, 9).Value = "Wire jumper"
            cell(1, 7).Value = XDV
            cell(1, 8).Value = "bk"
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
            Else
 End If
 End If
 
             If cell.Value = "XDV" And cell.Value = cell(1, 4).Value And cell(1, 2).Value = 11 And cell(1, 5).Value = 10 And Not (cell(1, 9).Value = "Wire jumper" Or cell(1, 9).Value = "Ponticello a filo") Then
            answer = MsgBox("Is the connection between XDV:10 and XDV:11  is with wire jumper ?", vbYesNo + vbQuestion, "-XDV jumpers")
            If answer = vbYes Then
            cell(1, 9).Value = "Wire jumper"
            cell(1, 7).Value = XDV
            cell(1, 8).Value = "bk"
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
End If
 End If
          
             If cell.Value = "XDV" And cell.Value = cell(1, 4).Value And cell(1, 2).Value = 12 And cell(1, 5).Value = 13 And Not (cell(1, 9).Value = "Wire jumper" Or cell(1, 9).Value = "Ponticello a filo") Then
            answer = MsgBox("Is the connection between XDV:12 and XDV:13  is with wire jumper ?", vbYesNo + vbQuestion, "-XDV jumpers")
            If answer = vbYes Then
            cell(1, 9).Value = "Wire jumper"
            cell(1, 7).Value = 1.5
            cell(1, 8).Value = "bk"
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
  End If
 End If
             If cell.Value = "XDV" And cell.Value = cell(1, 4).Value And cell(1, 2).Value = 13 And cell(1, 5).Value = 12 And Not (cell(1, 9).Value = "Wire jumper" Or cell(1, 9).Value = "Ponticello a filo") Then
            answer = MsgBox("Is the connection between XDV:12 and XDV:13  is with wire jumper ?", vbYesNo + vbQuestion, "-XDV jumpers")
            If answer = vbYes Then
            cell(1, 9).Value = "Wire jumper"
            cell(1, 7).Value = 1.5
            cell(1, 8).Value = "bk"
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
   
 End If
 End If
 Next
 '--------------------- - 'Ref protection- Connector 130----------------
 
    Set MyPlage = Range("A14:A1000")

    For Each cell In MyPlage

        If Left(cell.Value, 2) = "AA" And Left(cell(1, 2).Value, 5) = "-X130" And cell(1, 7).Value > 2.5 Then
            cell(1, 7).Interior.ColorIndex = 3
        End If
    
    
Next
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub


