Attribute VB_Name = "XDB1Connector"
Sub XDB1Connector()



'---------------------------XDB1----------------------------------------------

                   Set MyPlage = Range("D15:d1000")
        For Each cell In MyPlage
       
                       If cell.Value = "XDB1" And cell(1, 2).Value = 1 And cell(1, 4).Value < "2,5" Then
        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = "2,5"
        
        End If

                       If cell.Value = "XDB1" And cell(1, 2).Value = 25 And cell(1, 4).Value < "2,5" Then
        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = "2,5" '
        
        End If

        
           If cell.Value = "XDB1" And cell(1, 2).Value = 35 And cell(1, 4).Value < "2,5" Then
        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
       cell(1, 4).Value = "2,5"
        
        End If
                  

          If cell.Value = "XDB1" And cell(1, 2).Value = 40 And cell(1, 4).Value < "2,5" Then
        cell(1, 4).Font.ColorIndex = 3
        cell(1, 4).Font.Bold = True
        cell(1, 4).Value = "2,5"
        
        End If
                  

Next
            
     
     Set MyPlage = Range("A15:A1000")
        For Each cell In MyPlage
        
                      
                                   If cell.Value = "XDB1" And cell(1, 2).Value = 1 And cell(1, 7).Value < "2,5" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "2,5"
        
        End If


                       If cell.Value = "XDB1" And cell(1, 2).Value = 25 And cell(1, 7).Value < "2,5" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
       cell(1, 7).Value = "2,5"
        
        End If
                
        
        
           
                   If cell.Value = "XDB1" And cell(1, 2).Value = 35 And cell(1, 7).Value < "2,5" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "2,5"
        
        End If


          If cell.Value = "XDB1" And cell(1, 2).Value = 40 And cell(1, 7).Value < "2,5" Then
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        cell(1, 7).Value = "2,5"
        
        End If

 Next
      
          '-------------------------Connections"----------------------------------
   
    Set MyPlage = Range("D15:D1000")
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
   
       Set MyPlage = Range("A15:A1000")
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
    
    '-------------------------------Clear cells if have crossection------------------------------------------
    
    Set MyPlage = Range("A15:A1000")
        For Each cell In MyPlage
        
         If Not IsEmpty(cell(1, 7).Value) And cell.Value = "XDB1" And cell(1, 4).Value = "XDB" Then
            cell(1, 7).ClearContents
            cell(1, 8).ClearContents
            cell(1, 9).Value = "Direct connection"
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
            End If


Next
            
            
                
            
            
        '------------Swap Ranges----------
           Set MyPlage = Range("A15:A1000")
        For Each cell In MyPlage
        
 If cell.Value = "XDB" And cell(1, 4).Value = "XDB1" Then
  Dim FirstAddress As String
    Dim MyArr As Variant
    Dim Rng As Range
    Dim i As Long
    With Application
    'Dim s1 As String, s2 As String
    Dim r1 As Range, r2 As Range
    Dim temp1, temp2


        .ScreenUpdating = False
        .EnableEvents = False
    End With


    MyArr = Array("XDB")

    'Search Column or range
    With ActiveSheet.Columns("A")

        'clear the cells in the column to the right
        '.Offset(0, 1).ClearContents

        For i = LBound(MyArr) To UBound(MyArr)

            'If you want to find a part of the rng.value then use xlPart
            'if you use LookIn:=xlValues it will also work with a
            'formula cell that evaluates to "ron"

            Set Rng = .Find(What:=MyArr(i), _
                            After:=.Cells(.Cells.Count), _
                            LookIn:=xlFormulas, _
                            LookAt:=xlWhole, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlNext, _
                            MatchCase:=False)
  If Not Rng Is Nothing Then
                FirstAddress = Rng.Address
                Do

                    Set r1 = Rng
                    Set r2 = Rng

    temp1 = r1.Offset(, 0).Resize(, 3).Value
    temp2 = r2.Offset(, 3).Resize(, 3).Value


    r1.Offset(, 0).Resize(, 3).Value = temp2
    r2.Offset(, 3).Resize(, 3).Value = temp1



                    Set Rng = .FindNext(Rng)
                    If Rng Is Nothing Then
                        Exit Do
                    End If
           

                Loop While Rng.Address <> FirstAddress
                
            End If
        Next i
    End With
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With
    
     End If
 Next

    
End Sub
