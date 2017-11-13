Attribute VB_Name = "ERRORS_old"
Sub Errors_old()

Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
'------------------XDA -------------------------
'Range("G7:H1000").Interior.ColorIndex = 0
Set MyPlage = Range("A15:A1000")
  For Each cell In MyPlage
  
        
            If Left(cell.Value, 3) = "XDA" And cell(1, 7).Value = 1 Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = 4
        Else
        End If
               If Left(cell.Value, 3) = "XDA" And cell(1, 7).Value = 1.5 Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = 4
        Else
    
        End If
                     If Left(cell.Value, 3) = "XDA" And cell(1, 7).Value = 2.5 Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = 4
        Else
        End If
        
                       If Left(cell.Value, 3) = "XDA" And cell(1, 7).Value = "" Then
        cell(1, 7).Interior.ColorIndex = 16
        cell(1, 8).Interior.ColorIndex = 16
        Else
        
        End If
        
                    If Left(cell.Value, 3) = "XDA" And cell(1, 7).Value = "1" Then

                cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = 4
        Else
        End If
               If Left(cell.Value, 3) = "XDA" And cell(1, 7).Value = "1,5" Then
           cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = 4
        Else
        End If
                     If Left(cell.Value, 3) = "XDA" And cell(1, 7).Value = "2,5" Then
         cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = 4
        Else
        End If
        
    Next
    Set MyPlage = Range("D15:D1000")
  For Each cell In MyPlage
  
        
            If Left(cell.Value, 3) = "XDA" And cell(1, 4).Value = 1 Then

        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = 4
        
        End If
               If Left(cell.Value, 3) = "XDA" And cell(1, 4).Value = 1.5 Then
               cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = 4
        End If
                     If Left(cell.Value, 3) = "XDA" And cell(1, 4).Value = 2.5 Then
                cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = 4
        End If
        
                       If Left(cell.Value, 3) = "XDA" And cell(1, 4).Value = "" Then
        cell(1, 4).Interior.ColorIndex = 16
        cell(1, 5).Interior.ColorIndex = 16
        End If
        
                    If Left(cell.Value, 3) = "XDA" And cell(1, 4).Value = "1" Then
                cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = 4
        End If
               If Left(cell.Value, 3) = "XDA" And cell(1, 4).Value = "1,5" Then
                cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = 4
        End If
                     If Left(cell.Value, 3) = "XDA" And cell(1, 4).Value = "2,5" Then
                cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = 4
        End If
         Next
         
 
 '------------------XDV -------------------------

Set MyPlage = Range("A15:A1000")
  For Each cell In MyPlage
  
        
            If Left(cell.Value, 3) = "XDV" And cell(1, 7).Value = 1 Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "1,5"
      
        End If
               If Left(cell.Value, 3) = "XDV" And cell(1, 7).Value = "1" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "1,5"

        End If
        Next
    Set MyPlage = Range("D15:D1000")
  For Each cell In MyPlage
  
        
            If Left(cell.Value, 3) = "XDV" And cell(1, 4).Value = 1 Then
        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = "1,5"
         End If
                     If Left(cell.Value, 3) = "XDV" And cell(1, 4).Value = "1" Then
        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = "1,5"
         End If
         Next
         
     
        '---------------------------FCM----------------------------------------------
        Set MyPlage = Range("A15:A1000")
        For Each cell In MyPlage
        
If Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 3 And cell(1, 4).Value = "XDI2" And cell(1, 7).Value = "1" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "2,5"
        End If
        If Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 3 And cell(1, 4).Value = "XDI2" And cell(1, 7).Value = "1,5" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "2,5"
        End If
                If Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 3 And cell(1, 4).Value = "XDI2" And cell(1, 7).Value = "1.5" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "2,5"
        End If
        If Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 1 And cell(1, 4).Value = "XDI2" And cell(1, 7).Value = "1" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "2,5"
        End If
        If Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 1 And cell(1, 4).Value = "XDI2" And cell(1, 7).Value = "1,5" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "2,5"
        End If
                If Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 1 And cell(1, 4).Value = "XDI2" And cell(1, 7).Value = "1.5" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "2,5"
        End If
        
        If Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 3 And cell(1, 4).Value = "XDI3" And cell(1, 7).Value = "1" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "2,5"
        End If
        If Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 3 And cell(1, 4).Value = "XDI3" And cell(1, 7).Value = "1,5" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "2,5"
        End If
                If Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 3 And cell(1, 4).Value = "XDI3" And cell(1, 7).Value = "1.5" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "2,5"
        End If
        If Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 1 And cell(1, 4).Value = "XDI3" And cell(1, 7).Value = "1" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "2,5"
        End If
        If Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 1 And cell(1, 4).Value = "XDI3" And cell(1, 7).Value = "1,5" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "2,5"
        End If
                If Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 1 And cell(1, 4).Value = "XDI3" And cell(1, 7).Value = "1.5" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "2,5"
        End If
        If Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 3 And cell(1, 4).Value = "XDI4" And cell(1, 7).Value = "1" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "2,5"
        End If
        If Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 3 And cell(1, 4).Value = "XDI4" And cell(1, 7).Value = "1,5" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "2,5"
        End If
                If Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 3 And cell(1, 4).Value = "XDI4" And cell(1, 7).Value = "1.5" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "2,5"
        End If
        If Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 1 And cell(1, 4).Value = "XDI4" And cell(1, 7).Value = "1" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "2,5"
        End If
        If Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 1 And cell(1, 4).Value = "XDI4" And cell(1, 7).Value = "1,5" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "2,5"
        End If
                If Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 1 And cell(1, 4).Value = "XDI4" And cell(1, 7).Value = "1.5" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "2,5"
        End If
        If Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 3 And cell(1, 4).Value = "XDI5" And cell(1, 7).Value = "1" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "2,5"
        End If
        If Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 3 And cell(1, 4).Value = "XDI5" And cell(1, 7).Value = "1,5" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "2,5"
        End If
                If Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 3 And cell(1, 4).Value = "XDI5" And cell(1, 7).Value = "1.5" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "2,5"
        End If
        If Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 1 And cell(1, 4).Value = "XDI5" And cell(1, 7).Value = "1" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "2,5"
        End If
        If Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 1 And cell(1, 4).Value = "XDI5" And cell(1, 7).Value = "1,5" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "2,5"
        End If
                If Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 1 And cell(1, 4).Value = "XDI5" And cell(1, 7).Value = "1.5" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "2,5"
        End If
        If Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 3 And cell(1, 4).Value = "XDI7" And cell(1, 7).Value = "1" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "2,5"
        End If
        If Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 3 And cell(1, 4).Value = "XDI7" And cell(1, 7).Value = "1,5" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "2,5"
        End If
                If Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 3 And cell(1, 4).Value = "XDI7" And cell(1, 7).Value = "1.5" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "2,5"
        End If
        If Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 1 And cell(1, 4).Value = "XDI7" And cell(1, 7).Value = "1" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "2,5"
        End If
        If Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 1 And cell(1, 4).Value = "XDI7" And cell(1, 7).Value = "1,5" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "2,5"
        End If
                If Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 1 And cell(1, 4).Value = "XDI7" And cell(1, 7).Value = "1.5" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "2,5"
        End If
        If Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 3 And cell(1, 4).Value = "XDI8" And cell(1, 7).Value = "1" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "2,5"
        End If
        If Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 3 And cell(1, 4).Value = "XDI8" And cell(1, 7).Value = "1,5" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "2,5"
        End If
                If Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 3 And cell(1, 4).Value = "XDI8" And cell(1, 7).Value = "1.5" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "2,5"
        End If
        If Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 1 And cell(1, 4).Value = "XDI8" And cell(1, 7).Value = "1" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "2,5"
        End If
        'If Left(Cell.Value, 3) = "FCM" And Cell(1, 2).Value = 1 And Cell(1, 4).Value = "XDI8" And Cell(1, 7).Value = "1,5" Then
        'Cell(1, 7).Font.ColorIndex = 3
       ' Cell(1, 7).Font.Bold = True
        'Cell(1, 7).Value = "2,5"
        'End If
              '  If Left(Cell.Value, 3) = "FCM" And Cell(1, 2).Value = 1 And Cell(1, 4).Value = "XDI8" And Cell(1, 7).Value = "1.5" Then
        'Cell(1, 7).Font.ColorIndex = 3
        'Cell(1, 7).Font.Bold = True
        'Cell(1, 7).Value = "2,5"
        'End If
        If Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 3 And cell(1, 4).Value = "XDI9" And cell(1, 7).Value = "1" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "2,5"
        End If
        If Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 3 And cell(1, 4).Value = "XDI9" And cell(1, 7).Value = "1,5" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "2,5"
        End If
                If Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 3 And cell(1, 4).Value = "XDI9" And cell(1, 7).Value = "1.5" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "2,5"
        End If
        If Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 1 And cell(1, 4).Value = "XDI9" And cell(1, 7).Value = "1" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "2,5"
        End If
        If Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 1 And cell(1, 4).Value = "XDI9" And cell(1, 7).Value = "1,5" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "2,5"
        End If
                If Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 1 And cell(1, 4).Value = "XDI9" And cell(1, 7).Value = "1.5" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "2,5"
        End If

If Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 3 And cell(1, 4).Value = "XDI6" And cell(1, 7).Value = "1" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "1,5"
        End If
If Left(cell.Value, 3) = "FCM" And cell(1, 2).Value = 1 And cell(1, 4).Value = "XDI6" And cell(1, 7).Value = "1" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "1,5"
        End If

 Next
        
         
   '---------------------------XDI6----------------------------------------------
    
 Set MyPlage = Range("D15:d1000")
  For Each cell In MyPlage
  
        
            If cell.Value = "XDI6" And cell(1, 4).Value = 1 Then
        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = "1,5"
        
        End If
                   If cell.Value = "XDI6" And cell(1, 4).Value = "1" Then
        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = "1,5"
                  
           End If
         Next
         
            '---------------------------XDI8----------------------------------------------
    Set MyPlage = Range("A15:A1000")
  For Each cell In MyPlage
  
        
            If cell.Value = "XDI8" And cell(1, 7).Value = 1 Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = 4
        Else
        End If
               If cell.Value = "XDI8" And cell(1, 7).Value = 1.5 Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = 4
        Else
    
        End If
                     If cell.Value = "XDI8" And cell(1, 7).Value = 2.5 Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = 4
        Else
        End If
        
        
                    If cell.Value = "XDI8" And cell(1, 7).Value = "1" Then

                cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = 4
        Else
        End If
               If cell.Value = "XDI8" And cell(1, 7).Value = "1,5" Then
           cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = 4
        Else
        End If
                     If cell.Value = "XDI8" And cell(1, 7).Value = "2,5" Then
         cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = 4
        Else
        End If
        
    Next
    
    
    Set MyPlage = Range("D15:D1000")
  For Each cell In MyPlage
  
        
            If cell.Value = "XDI8" And cell(1, 4).Value = 1 Then

        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = 4
        
        End If
               If cell.Value = "XDI8" And cell(1, 4).Value = 1.5 Then
               cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = 4
        End If
                     If cell.Value = "XDI8" And cell(1, 4).Value = 2.5 Then
                cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = 4
        End If

        
                    If cell.Value = "XDI8" And cell(1, 4).Value = "1" Then
                cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = 4
        End If
               If cell.Value = "XDI8" And cell(1, 4).Value = "1,5" Then
                cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = 4
        End If
                     If cell.Value = "XDI8" And cell(1, 4).Value = "2,5" Then
                cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = 4
        End If
         Next
   '---------------------------XDI2----------------------------------------------
  Set MyPlage = Range("A15:A1000")
  For Each cell In MyPlage
  
        
            If cell.Value = "XDI2" And cell(1, 7).Value = 1 Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "2,5"
        Else
        End If
               If cell.Value = "XDI2" And cell(1, 7).Value = 1.5 Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "2,5"
        Else
        
        End If
        
                    If cell.Value = "XDI2" And cell(1, 7).Value = "1" Then

                cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "2,5"
        Else
        End If
               If cell.Value = "XDI2" And cell(1, 7).Value = "1,5" Then
           cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "2,5"
        Else
        End If

        
    Next
    Set MyPlage = Range("D15:D1000")
  For Each cell In MyPlage
  
        
            If cell.Value = "XDI2" And cell(1, 4).Value = 1 Then

        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = "2,5"
        
        End If
               If cell.Value = "XDI2" And cell(1, 4).Value = 1.5 Then
               cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = "2,5"
        End If

        
                    If cell.Value = "XDI2" And cell(1, 4).Value = "1" Then
                cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = "2,5"
        End If
               If cell.Value = "XDI2" And cell(1, 4).Value = "1,5" Then
                cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = "2,5"
        End If

         Next
         
         
      '---------------------------XDI3----------------------------------------------
Set MyPlage = Range("A15:A1000")
  For Each cell In MyPlage
  
        
            If cell.Value = "XDI3" And cell(1, 7).Value = 1 Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "2,5"
        Else
        End If
               If cell.Value = "XDI3" And cell(1, 7).Value = 1.5 Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "2,5"
        Else
        
        End If
        
                    If cell.Value = "XDI3" And cell(1, 7).Value = "1" Then

                cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "2,5"
        Else
        End If
               If cell.Value = "XDI3" And cell(1, 7).Value = "1,5" Then
           cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "2,5"
        Else
        End If

        
    Next
    Set MyPlage = Range("D15:D1000")
  For Each cell In MyPlage
  
        
            If cell.Value = "XDI3" And cell(1, 4).Value = 1 Then

        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = "2,5"
        
        End If
               If cell.Value = "XDI3" And cell(1, 4).Value = 1.5 Then
               cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = "2,5"
        End If

        
                    If cell.Value = "XDI3" And cell(1, 4).Value = "1" Then
                cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = "2,5"
        End If
               If cell.Value = "XDI3" And cell(1, 4).Value = "1,5" Then
                cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = "2,5"
        End If

         Next
         
      '--------------------------------PGA--------------------------------------------


Set MyPlage = Range("A15:A1000")
  For Each cell In MyPlage
  
        
            If Left(cell.Value, 3) = "PGA" And cell(1, 7).Value = 1 Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = 4
        End If
        
               If Left(cell.Value, 3) = "PGA" And cell(1, 7).Value = 1.5 Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = 4
        End If
                       If Left(cell.Value, 3) = "PGA" And cell(1, 7).Value = 2.5 Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = 4
        End If
        
                    If Left(cell.Value, 3) = "PGA" And cell(1, 7).Value = "1" Then

                cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = 4
        End If
               If Left(cell.Value, 3) = "PGA" And cell(1, 7).Value = "1,5" Then
           cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = 4
        End If
               If Left(cell.Value, 3) = "PGA" And cell(1, 7).Value = "2,5" Then
           cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = 4
        End If
        
    Next
    Set MyPlage = Range("D15:D1000")
  For Each cell In MyPlage
  
        
            If Left(cell.Value, 3) = "PGA" And cell(1, 4).Value = 1 Then

        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = 4
        
        End If
               If Left(cell.Value, 3) = "PGA" And cell(1, 4).Value = 1.5 Then
               cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = 4
        End If
               If Left(cell.Value, 3) = "PGA" And cell(1, 4).Value = 2.5 Then
               cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = 4
        End If
        
                    If Left(cell.Value, 3) = "PGA" And cell(1, 4).Value = "1" Then
                cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = 4
        End If
               If Left(cell.Value, 3) = "PGA" And cell(1, 4).Value = "1,5" Then
                cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = 4
        End If
               If Left(cell.Value, 3) = "PGA" And cell(1, 4).Value = "2,5" Then
                cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = 4
        End If

   Next
   
   
   '--------------------------------PGV--------------------------------------------


Set MyPlage = Range("A15:A1000")
  For Each cell In MyPlage
  
        
            If Left(cell.Value, 3) = "PGV" And cell(1, 7).Value = 1 Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "1,5"
        End If
        

            If Left(cell.Value, 3) = "PGV" And cell(1, 7).Value = "1" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "1,5"
        End If


        
    Next
    Set MyPlage = Range("D15:D1000")
  For Each cell In MyPlage
  
        
            If Left(cell.Value, 3) = "PGV" And cell(1, 4).Value = 1 Then

        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = "1,5"
        
        End If

                    If Left(cell.Value, 3) = "PGV" And cell(1, 4).Value = "1" Then
                cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = "1,5"
        End If
        Next
'--------------------------------XDM1--------------------------------------------


Set MyPlage = Range("A15:A1000")
  For Each cell In MyPlage
  
        
            If cell.Value = "XDM1" And cell(1, 7).Value = 1 Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = 4
        End If
        
               If cell.Value = "XDM1" And cell(1, 7).Value = 1.5 Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = 4
        End If
                       If cell.Value = "XDM1" And cell(1, 7).Value = 2.5 Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = 4
        End If
        
                    If cell.Value = "XDM1" And cell(1, 7).Value = "1" Then

                cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = 4
        End If
               If cell.Value = "XDM1" And cell(1, 7).Value = "1,5" Then
           cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = 4
        End If
               If cell.Value = "XDM1" And cell(1, 7).Value = "2,5" Then
           cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = 4
        End If
        
    Next
    Set MyPlage = Range("D15:D1000")
  For Each cell In MyPlage
  
        
            If cell.Value = "XDM1" And cell(1, 4).Value = 1 Then

        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = 4
        
        End If
               If cell.Value = "XDM1" And cell(1, 4).Value = 1.5 Then
               cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = 4
        End If
               If cell.Value = "XDM1" And cell(1, 4).Value = 2.5 Then
               cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = 4
        End If
        
                    If cell.Value = "XDM1" And cell(1, 4).Value = "1" Then
                cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = 4
        End If
               If cell.Value = "XDM1" And cell(1, 4).Value = "1,5" Then
                cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = 4
        End If
               If cell.Value = "XDM1" And cell(1, 4).Value = "2,5" Then
                cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = 4
        End If

   Next
   
   
   '--------------------------------XDM2--------------------------------------------


Set MyPlage = Range("A15:A1000")
  For Each cell In MyPlage
  
        
            If cell.Value = "XDM2" And cell(1, 7).Value = 1 Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "1,5"
        End If
        

            If cell.Value = "XDM2" And cell(1, 7).Value = "1" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "1,5"
        End If


        
    Next
    Set MyPlage = Range("D15:D1000")
  For Each cell In MyPlage
  
        
            If cell.Value = "XDM2" And cell(1, 4).Value = 1 Then

        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = "1,5"
        
        End If

                    If cell.Value = "XDM2" And cell(1, 4).Value = "1" Then
                cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = "1,5"
        End If
   Next
   
      '-------------------------XE"----------------------------------
   
    Set MyPlage = Range("D15:D1000")
  For Each cell In MyPlage
  
        
        If Not (cell(1, 5).Value = "gnye" Or cell(1, 5).Value = "GNYE") And Left(cell.Value, 2) = "XE" Then
        cell(1, 5).Value = "gnye"
        cell(1, 5).Font.ColorIndex = 3
        cell(1, 5).Font.Bold = True
        End If
          'If cell(1, 5).Value = "GNYE" And Left(cell.Value, 2) = "XE" Then
        'cell(1, 5).Interior.ColorIndex = 0
        'cell(1, 5).Font.ColorIndex = 0
        'End If
        
         If Left(cell.Value, 2) = "XE" And cell(1, 4).Value = "1" Then
        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = "2,5"
        End If
        
                 If Left(cell.Value, 2) = "XE" And cell(1, 4).Value = 1 Then
        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = "2,5"
        End If
                 If Left(cell.Value, 2) = "XE" And cell(1, 4).Value = "1,5" Then
        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = "2,5"
        End If
        
                 If Left(cell.Value, 2) = "XE" And cell(1, 4).Value = 1.5 Then
        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = "2,5"
        End If
        
   Next
   
   '------------------------- Jumpers clear cells----------------------------------
   
   Jumpers.Jumpers
   
   
   '------------------XDB93 -XDB91----------------------------------
Set MyPlage = Range("A15:A1000")
  For Each cell In MyPlage
  
        
            If cell.Value = "XDB91" Then
answer = MsgBox("Please check for other connection between XDB93 and XDB91!!!", vbYes + vbQuestion, "")
 End If

          
Next
Set MyPlage = Range("D15:D1000")
  For Each cell In MyPlage
  
        
            If cell.Value = "XDB91" Then
answer = MsgBox("Please check for other connection between XDB93 and XDB91!!!", vbYes + vbQuestion, "")
 End If

          
Next

 

Set MyPlage = Range("D15:D1000")
  For Each cell In MyPlage
  
        
            If cell.Value = "XDB91" Then
answer = MsgBox("Please check for other connection between XDB93 and XDB91!!!", vbYes + vbQuestion, "")
 End If

          
Next

 '------------------XDV -------------------------
Set MyPlage = Range("A15:A1000")
  For Each cell In MyPlage
  
        
            If cell.Value = "XDV" And cell.Value = cell(1, 4).Value And cell(1, 2).Value = 10 And cell(1, 5).Value = 11 Then
            cell(1, 9).Value = "Wire jumper"
            cell(1, 7).Value = "1,5"
            cell(1, 8).Value = "bk"
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
            
   answer = MsgBox("Please check connection XDV:10 to XDV:11 if not wire jumper then remove ection and colour!!!", vbYes + vbQuestion, "")
 End If
             If cell.Value = "XDV" And cell.Value = cell(1, 4).Value And cell(1, 2).Value = 11 And cell(1, 5).Value = 10 Then
            cell(1, 9).Value = "Wire jumper"
            cell(1, 7).Value = "1,5"
            cell(1, 8).Value = "bk"
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
   answer = MsgBox("Please check connection XDV:10 to XDV:11 if not wire jumper then remove ection and colour!!!", vbYes + vbQuestion, "")
 End If
          
             If cell.Value = "XDV" And cell.Value = cell(1, 4).Value And cell(1, 2).Value = 12 And cell(1, 5).Value = 13 Then
            cell(1, 9).Value = "Wire jumper"
            cell(1, 7).Value = "1,5"
            cell(1, 8).Value = "bk"
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
   answer = MsgBox("Please check connection XDV:10 to XDV:11 if not wire jumper then remove ection and colour!!!", vbYes + vbQuestion, "")
 End If
             If cell.Value = "XDV" And cell.Value = cell(1, 4).Value And cell(1, 2).Value = 13 And cell(1, 5).Value = 12 Then
            cell(1, 9).Value = "Wire jumper"
            cell(1, 7).Value = "1,5"
            cell(1, 8).Value = "bk"
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
   answer = MsgBox("Please check connection XDV:10 to XDV:11 if not wire jumper then remove ection and colour!!!", vbYes + vbQuestion, "")
 End If
                    
          
Next

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub


