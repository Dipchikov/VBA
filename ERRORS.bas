Attribute VB_Name = "ERRORS"

Public Sub Errors()
Dim lr As Long
lr = Range("A" & Rows.Count).End(xlUp).Row
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False

Dim XDA As Double
Dim xdv As Double
Dim XDB1 As Double
Dim FCMm As Double
'----------default values----------
XDA = XDA1
xdv = XDV1
XDB1 = motor
FCMm = XDB1

         
'------------------XDV -terminal------------------------

Set MyPlage = Range("A15:A" & lr)
  For Each cell In MyPlage
  
        
            If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "XDV" And Not (Left(cell(1, 4).Value, 2) = "XE" Or Left(cell(1, 4).Value, 2) = "PE") And cell(1, 7).Value <> xdv Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = xdv
      
        End If
              
        Next
    Set MyPlage = Range("D15:D" & lr)
  For Each cell In MyPlage
  
        
            If Not IsEmpty(cell(1, 4).Value) And Left(cell.Value, 3) = "XDV" And cell(1, 4).Value <> xdv Then
        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = xdv
        End If
                     
 Next
         

 
 

'------------------XDV - connections----------
Set MyPlage = Range("A15:A" & lr)
  For Each cell In MyPlage
  
        
            If cell.Value = "XDV" And cell.Value = cell(1, 4).Value Then
            answer = MsgBox("Is connection between" & cell(1, 3).Value & " And " & cell(1, 6).Value & " is with-" & cell(1, 9).Value, vbYesNo + vbQuestion, "-XDV jumpers")
            If answer = vbYes And (cell(1, 9).Value = "Saddle jumper" Or cell(1, 9).Value = "Insertable jumper") Then
            cell(1, 7).ClearContents
            cell(1, 8).ClearContents
            If cell(1, 9).Value = "Saddle jumper" Then
            cell(1, 9).Value = "Insertable jumper"
            cell(1, 9).Font.ColorIndex = 6
            cell(1, 9).Font.Bold = True
            End If
            End If
            If answer = vbNo Then
            If Not (cell(1, 9).Value = "Wire jumper") Then
            cell(1, 9).Value = "Wire jumper"
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
            End If
            If IsEmpty(cell(1, 7).Value) Then
            cell(1, 7).Value = xdv
            cell(1, 7).Font.ColorIndex = 3
            cell(1, 7).Font.Bold = True
            End If
            If IsEmpty(cell(1, 8).Value) Then
            cell(1, 8).Value = "bk"
            cell(1, 8).Font.ColorIndex = 3
            cell(1, 8).Font.Bold = True
            End If
            
           If answer = vbYes And (cell(1, 9).Value = "Wire jumper" Or cell(1, 9).Value = "Conductor / wire") Then
            If Not (cell(1, 9).Value = "Wire jumper") Then
            cell(1, 9).Value = "Wire jumper"
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
            End If
            If IsEmpty(cell(1, 7).Value) Then
            cell(1, 7).Value = xdv
            cell(1, 7).Font.ColorIndex = 3
            cell(1, 7).Font.Bold = True
            End If
            If IsEmpty(cell(1, 8).Value) Then
            cell(1, 8).Value = "bk"
            cell(1, 8).Font.ColorIndex = 3
            cell(1, 8).Font.Bold = True
            End If
            
            If answer = vbNo And Not (cell(1, 9).Value = "Insertable jumper") Then
            cell(1, 7).ClearContents
            cell(1, 8).ClearContents
            cell(1, 9).Value = "Insertable jumper"
            cell(1, 9).Font.ColorIndex = 6
            cell(1, 9).Font.Bold = True
            End If


End If
End If
End If
         
Next

         
'------------------XDA -------------------------
'Range("G7:H1000").Interior.ColorIndex = 0
Set MyPlage = Range("A15:A" & lr)
  For Each cell In MyPlage
  
        
    If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "XDA" And cell(1, 7).Value <> XDA Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = XDA
        End If
        
    Next
    Set MyPlage = Range("D15:D" & lr)
  For Each cell In MyPlage
  
        
            If Not IsEmpty(cell(1, 4).Value) And Left(cell.Value, 3) = "XDA" And cell(1, 4).Value <> XDA Then
        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = XDA
        End If


 Next
         
     
        '---------------------------FCM----------------------------------------------
        Set MyPlage = Range("A15:A" & lr)
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
        
         Set MyPlage = Range("A15:A" & lr)
        For Each cell In MyPlage
        
        If Not IsEmpty(cell(1, 7).Value) And cell.Value = "XDI6" And cell(1, 7).Value <> xdv Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = xdv
        End If
        Next
 
    'If Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 3 And cell(1, 4).Value = "XDI6" And cell(1, 7).Value <> XDV Then
       ' cell(1, 7).Font.ColorIndex = 3
       '' cell(1, 7).Font.Bold = True
       ' cell(1, 7).Value = XDV
       ' End If
        'Next
        
    
 Set MyPlage = Range("D15:D" & lr)
  For Each cell In MyPlage
  
        
If Not IsEmpty(cell(1, 4).Value) And cell.Value = "XDI6" And cell(1, 4).Value <> xdv Then
        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = xdv
        
 End If
                
Next

'---------------------------XDI7----------------------------------------------
        
        Set MyPlage = Range("A15:A" & lr)
        For Each cell In MyPlage
        
If cell.Value = "XDI7" And cell(1, 12).Value <> "Shielded cable" Then
answer = MsgBox("Is connection between" & cell(1, 3).Value & " And " & cell(1, 6).Value & " is with Shielded cable", vbYesNo + vbQuestion, "-XDI7 Shielded cable")  'is with " & cell(1, 9).Value
        If answer = vbYes Then
        cell(1, 12).Value = "Shielded cable"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
 
          
If cell(1, 4).Value = "XDI7" And cell(1, 12).Value <> "Shielded cable" Then
answer = MsgBox("Is connection between" & cell(1, 3).Value & " And " & cell(1, 6).Value & " is with Shielded cable", vbYesNo + vbQuestion, "-XDI7 Shielded cable") 'is with " & cell(1, 9).Value
If answer = vbYes Then

        cell(1, 12).Value = "Shielded cable"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
 End If
 End If
                
Next


         
'---------------------------XDI8----------------------------------------------
    Set MyPlage = Range("A15:A" & lr)
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
  'Set MyPlage = Range("A15:A" & lr)
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
   
  Set MyPlage = Range("A15:A" & lr)
  For Each cell In MyPlage
  
        
            If Not IsEmpty(cell(1, 7).Value) And cell.Value = "XDI2" And cell(1, 7).Value <> XDB1 Then
            If Not Left(cell(1, 5).Value, 1) = "A" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = XDB1
        End If
        End If
        
             'If Not IsEmpty(cell(1, 7).Value) And cell.Value = "XDI2" And cell(1, 7).Value < 2.5 And cell(1, 13).Value >= 2 Then
        'cell(1, 7).Font.ColorIndex = 3
        'cell(1, 7).Font.Bold = True
        'cell(1, 7).Value = 2.5
       ' End If
        '

        
             If Not IsEmpty(cell(1, 7).Value) And cell(1, 4).Value = "XDI2" And cell(1, 7).Value <> XDB1 Then
            If Not Left(cell(1, 2).Value, 1) = "A" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = XDB1
        End If
        End If
        
        ' If Not IsEmpty(cell(1, 7).Value) And cell(1, 4).Value = "XDI2" And cell(1, 7).Value < 2.5 And cell(1, 14).Value >= 2 Then
       ' cell(1, 7).Font.ColorIndex = 3
       ' cell(1, 7).Font.Bold = True
       ' cell(1, 7).Value = 2.5
       ' End If
 
    Next


         
         
      '---------------------------XDI3----------------------------------------------
Set MyPlage = Range("A15:A" & lr)
  For Each cell In MyPlage
  
        
            If Not IsEmpty(cell(1, 7).Value) And cell.Value = "XDI3" And cell(1, 7).Value <> XDB1 Then
            If Not Left(cell(1, 5).Value, 1) = "A" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = XDB1
        End If
        End If
        
         ' If Not IsEmpty(cell(1, 7).Value) And cell.Value = "XDI3" And cell(1, 7).Value < 2.5 And cell(1, 13).Value >= 2 Then
       ' cell(1, 7).Font.ColorIndex = 3
       ' cell(1, 7).Font.Bold = True
       ' cell(1, 7).Value = 2.5
      '  End If
        
             If Not IsEmpty(cell(1, 7).Value) And cell(1, 4).Value = "XDI3" And cell(1, 7).Value <> XDB1 Then
            If Not Left(cell(1, 2).Value, 1) = "A" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = XDB1
        End If
        End If
    'If Not IsEmpty(cell(1, 7).Value) And cell(1, 4).Value = "XDI3" And cell(1, 7).Value < 2.5 And cell(1, 14).Value >= 2 Then
        'cell(1, 7).Font.ColorIndex = 3
        'cell(1, 7).Font.Bold = True
        'cell(1, 7).Value = 2.5
        'End If

Next
         
      '--------------------------------PGA--------------------------------------------


Set MyPlage = Range("A15:A" & lr)
  For Each cell In MyPlage
  
        
            If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "PGA" And cell(1, 7).Value <> XDA Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = XDA
        End If
                
 Next
 
    Set MyPlage = Range("D15:D" & lr)
  For Each cell In MyPlage
  
        
            If Not IsEmpty(cell(1, 4).Value) And Left(cell.Value, 3) = "PGA" And cell(1, 4).Value <> XDA Then
        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = XDA
        
        End If
               
Next
   
   
   '--------------------------------PGV--------------------------------------------


Set MyPlage = Range("A15:A" & lr)
  For Each cell In MyPlage
  
        
            If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "PGV" And cell(1, 7).Value <> xdv Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = xdv
        End If
Next
        
    Set MyPlage = Range("D15:D" & lr)
  For Each cell In MyPlage
  
        
            If Not IsEmpty(cell(1, 4).Value) And Left(cell.Value, 3) = "PGV" And cell(1, 4).Value <> xdv Then

        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = xdv
        
        End If

 Next
 

'--------------------------------XDMs-Eroors-------------------------------------------

'XDMs_errors.XDMs_errors
Dim XDM As Double 'Single


For i = 1 To 10
Set MyPlage = Range("A15:d1000")
Application.ScreenUpdating = False
lr = Range("A" & Rows.Count).End(xlUp).Row


Set myCell = MyPlage.Find(What:="XDM" & i, LookIn:=xlValues)
If Not myCell Is Nothing Then
XDM = InputBox("Please add cross-section of conductors" & vbNewLine & "If XDM" & i & " is from  current circuit usually the cross-section of conductors is = 4mm" & vbNewLine & "If XDM" & i & " is  from  voltage circuit usually the cross-section of conductors is = 1,5mm", "Cross-Section for XDM" & i, Range("G" & myCell.Row).Value)
End If
Set MyPlage = Range("A15:A" & lr)
  For Each cell In MyPlage
 
              If cell.Value = "XDM" & i And cell(1, 7).Value <> XDM Then ' Not IsEmpty(cell(1, 7).Value) And
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = XDM
        End If

 Next

 
  Set MyPlage = Range("D15:D" & lr)
  For Each cell In MyPlage
       
            If cell.Value = "XDM" & i And cell(1, 4).Value <> XDM Then  'Not IsEmpty(cell(1, 4).Value) And
        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = XDM
        End If
        
Next
Next i






'-------------------------XE"----------------------------------

 Set MyPlage = Range("A15:A" & lr)
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


   
    Set MyPlage = Range("D15:D" & lr)
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
Set MyPlage = Range("A15:A" & lr)
  For Each cell In MyPlage
  
        
If cell.Value = "XDB91" Then
answer = MsgBox("Please check for other connection between XDB93 and XDB91!!!", vbOKOnly + vbExclamation, "Connection between XDB93 and XDB91")
 End If

          
Next
Set MyPlage = Range("D15:D" & lr)
  For Each cell In MyPlage
  
 If cell.Value = "XDB91" Then
answer = MsgBox("Please check for other connection between XDB93 and XDB91!!!", vbOKOnly + vbExclamation, "Connection between XDB93 and XDB91")
 End If

          
Next




 '--------------------- - 'Ref protection- Connector 130----------------
 
    Set MyPlage = Range("A14:A1000")

    For Each cell In MyPlage

        If Left(cell.Value, 2) = "AA" And Left(cell(1, 2).Value, 5) = "-X130" And cell(1, 7).Value > 2.5 Then
            cell(1, 7).Interior.ColorIndex = 3
        End If
    
    
Next

'---------------------------XDB1 connector---------------------------------------------
'---------------------------------------------------------------------------------------------
If Error_menu.CheckBox2.Value = True And Error_menu.CheckBox1.Value = False Then

 Set MyPlage = Range("D15:d" & lr)
        For Each cell In MyPlage
       
            If cell.Value = "XDB1" And cell(1, 2).Value = 1 And cell(1, 4).Value <> XDB1 Then
        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = XDB1
        
        End If

        If cell.Value = "XDB1" And cell(1, 2).Value = 25 And cell(1, 4).Value <> XDB1 Then
        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = XDB1 '
        
        End If

        
           If cell.Value = "XDB1" And cell(1, 2).Value = 35 And cell(1, 4).Value <> XDB1 Then
        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
       cell(1, 4).Value = XDB1
        
        End If
                  

          If cell.Value = "XDB1" And cell(1, 2).Value = 40 And cell(1, 4).Value <> XDB1 Then
        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = XDB1
        
        End If
                  

Next
            
     
     Set MyPlage = Range("A15:A" & lr)
        For Each cell In MyPlage
        
                      
           If cell.Value = "XDB1" And cell(1, 2).Value = 1 And cell(1, 7).Value <> XDB1 Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = XDB1
        
        End If


          If cell.Value = "XDB1" And cell(1, 2).Value = 25 And cell(1, 7).Value <> XDB1 Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
       cell(1, 7).Value = XDB1
        
        End If
                
        
        
           
           If cell.Value = "XDB1" And cell(1, 2).Value = 35 And cell(1, 7).Value <> XDB1 Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = XDB1
        
        End If


          If cell.Value = "XDB1" And cell(1, 2).Value = 40 And cell(1, 7).Value <> XDB1 Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = XDB1
        
        End If

 Next
 
 
'-------------------------Connections"----------------------------------
      
 'XDB1_connectors_number.XDB1_connectors_number
    
 '-------------------------------Clear cells if have crossection------------------------------------------
    
    Set MyPlage = Range("A15:A" & lr)
        For Each cell In MyPlage
        
         If Not IsEmpty(cell(1, 7).Value) And cell.Value = "XDB1" And Not (cell(1, 4).Value = "XDB1") And Left(cell(1, 4).Value, 3) = "XDB" Then
            cell(1, 7).ClearContents
            cell(1, 8).ClearContents
            cell(1, 9).Value = "Direct connection"
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
            End If
Next
End If

'---------------------------XDB1 ado1---------------------------------------------
'---------------------------------------------------------------------------------------------

If Error_menu.CheckBox2.Value = False And Error_menu.CheckBox1.Value = True Then


'---------------------------XDB1----------------------------------------------
        Set MyPlage = Range("D15:d" & lr)
        For Each cell In MyPlage
       
                       If cell.Value = "XDB1" And cell(1, 2).Value = 1 And cell(1, 4).Value <> XDB1 Then
        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = XDB1
        
        End If

                   If cell.Value = "XDB1" And cell(1, 2).Value = 2 And cell(1, 4).Value <> XDB1 Then
        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = XDB1
        End If
        
           
 Next
            
  
     Set MyPlage = Range("A15:A" & lr)
        For Each cell In MyPlage
        
                      
        If cell.Value = "XDB1" And cell(1, 2).Value = 1 And cell(1, 7).Value <> XDB1 Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = XDB1
        
        End If
                   If cell.Value = "XDB1" And cell(1, 2).Value = 2 And cell(1, 7).Value <> XDB1 Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = XDB1
                  
           End If
           
          
    Next
            
     '---------------------------XDB----------------------------------------------
    
 Set MyPlage = Range("D15:d" & lr)
  For Each cell In MyPlage
  
        
            If cell.Value = "XDB" And cell(1, 2).Value = 1 And cell(1, 4).Value <> XDB1 Then
        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = XDB1
        
        End If
                   If cell.Value = "XDB" And cell(1, 2).Value = 2 And cell(1, 4).Value <> XDB1 Then
        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = XDB1
                  
         End If
Next



'---------------------------XDB1-XDB is Direct connection---------------------------------------------
         Set MyPlage = Range("A15:A" & lr)
        For Each cell In MyPlage
       
         If cell.Value = "XDB1" And cell(1, 4).Value = "XDB" And cell(1, 9).Value = "Direct connection" Then
        cell(1, 9).Value = "Conductor / wire"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
         End If
Next
      '-------------------------Connections"----------------------------------
      
   
        'XDB1ado_connections.XDB1ado_connections
        
End If



    '---------------------------FCM3----------------------------------------------
If Error_menu.CheckBox6.Value = True Then


  'FCMm = InputBox("Please add cross-section of conductors FCMm circuit" & vbNewLine & "Cross-section of conductors for FCMm circuit  by default is = 2,5", "Cross-Section for FCMm circuit", "2,5")
  Set MyPlage = Range("A15:A" & lr)
  For Each cell In MyPlage
  

            If cell.Value = "FCM3" And cell(1, 2).Value = 1 And cell(1, 7).Value <> FCMm Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = FCMm
        End If
                    
                    If cell.Value = "FCM3" And cell(1, 2).Value = 3 And cell(1, 7).Value <> FCMm Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = FCMm
        End If
                   
                    If cell.Value = "FCM3" And cell(1, 2).Value = 2 And cell(1, 7).Value <> FCMm Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = FCMm
        End If
                    
                    If cell.Value = "FCM3" And cell(1, 2).Value = 4 And cell(1, 7).Value <> FCMm Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = FCMm
        End If
                   
 Next
 End If
        
        


    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub


