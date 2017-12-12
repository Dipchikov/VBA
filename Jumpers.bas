Attribute VB_Name = "Jumpers"
Sub Jumpers()

Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False

 '------------------Jumpers Primery -------------------------
 
Set myplage = Range("G15:G1000")

    For Each cell In myplage

   If cell.Value = "Bridge" Then
            cell(1, 3).Value = "Insertable jumper"
            cell.Value = ""
    End If
    Next

'----------------- minimal wires crossection   --------------------
Dim wire As String

Set myplage = Range("G15:G1000")


wire = InputBox("Please add minimal cross-section of conductors", "Read the General Arrangement Drawings", 1)
For Each cell In myplage
 If Not IsEmpty(cell.Value) And cell.Value < wire Then
 If Not (cell(1, 6).Value = "-" Or cell(1, 6).Value = "Shielded cable") Then
 cell.Value = wire
 cell.Font.ColorIndex = 3
 cell.Font.Bold = True
End If
End If

Next




'---------------------------clear cells BAT-FCF -QAB -BGT -QCE -BCT- BCN- BAD- EB- EA-----------------------------------

Set myplage = Range("A15:A1000")

    For Each cell In myplage

        If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "BAT" Then
            cell(1, 7).ClearContents
            cell(1, 8).ClearContents
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
          End If

                If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "QCE" Then
            cell(1, 7).ClearContents
            cell(1, 8).ClearContents
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
          End If
          
        If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "FCF" Then
            cell(1, 7).ClearContents
            cell(1, 8).ClearContents
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True

        End If
        
        If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "QAB" Then
            cell(1, 7).ClearContents
            cell(1, 8).ClearContents
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
        End If
    
        If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "BGT" Then
            cell(1, 7).ClearContents
            cell(1, 8).ClearContents
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
        End If
        
                If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "BGE" Then
            cell(1, 7).ClearContents
            cell(1, 8).ClearContents
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
        End If
        
                If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "BCT" Then
            cell(1, 7).ClearContents
            cell(1, 8).ClearContents
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
        End If
               If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "BCN" Then
            cell(1, 7).ClearContents
            cell(1, 8).ClearContents
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
        End If
                 If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "BAD" Then
            cell(1, 7).ClearContents
            cell(1, 8).ClearContents
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
        End If
        
              'If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 2) = "EB" Then
           ' cell(1, 7).ClearContents
           ' cell(1, 8).ClearContents
           ' cell(1, 9).Font.ColorIndex = 3
           ' cell(1, 9).Font.Bold = True
       ' End If
        
                    '  If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 2) = "EA" Then
          '  cell(1, 7).ClearContents
          '  cell(1, 8).ClearContents
           ' cell(1, 9).Font.ColorIndex = 3
          ' cell(1, 9).Font.Bold = True
       ' End If
        
       Next
                 
                 
 Set myplage = Range("D15:D1000")

    For Each cell In myplage
    

          If Not IsEmpty(cell(1, 4).Value) And Left(cell.Value, 3) = "BAT" Then
          
            cell(1, 4).ClearContents
            cell(1, 5).ClearContents
            cell(1, 6).Font.ColorIndex = 3
            cell(1, 6).Font.Bold = True

        End If
        
        If Not IsEmpty(cell(1, 4).Value) And Left(cell.Value, 3) = "FCF" Then
            cell(1, 4).ClearContents
            cell(1, 5).ClearContents
            cell(1, 6).Font.ColorIndex = 3
            cell(1, 6).Font.Bold = True
        End If
            If Not IsEmpty(cell(1, 4).Value) And Left(cell.Value, 3) = "QAB" Then
            cell(1, 4).ClearContents
            cell(1, 5).ClearContents
            cell(1, 6).Font.ColorIndex = 3
            cell(1, 6).Font.Bold = True
        End If

            If Not IsEmpty(cell(1, 4).Value) And Left(cell.Value, 3) = "BGT" Then
            cell(1, 4).ClearContents
            cell(1, 5).ClearContents
            cell(1, 6).Font.ColorIndex = 3
            cell(1, 6).Font.Bold = True
        End If
            If Not IsEmpty(cell(1, 4).Value) And Left(cell.Value, 3) = "BGE" Then
            cell(1, 4).ClearContents
            cell(1, 5).ClearContents
            cell(1, 6).Font.ColorIndex = 3
            cell(1, 6).Font.Bold = True
        End If
                    If Not IsEmpty(cell(1, 4).Value) And Left(cell.Value, 3) = "QCE" Then
            cell(1, 4).ClearContents
            cell(1, 5).ClearContents
            cell(1, 6).Font.ColorIndex = 3
            cell(1, 6).Font.Bold = True
        End If
        
        Next
 '------------------Jumpers between equipment -------------------------
 
Set myplage = Range("A15:A1000")

    For Each cell In myplage

   If cell.Value <> cell(1, 4).Value And cell(1, 9).Value = "Ponticello a staffa" Then
            cell(1, 9).Value = "Conduttore/filo"
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
            If IsEmpty(cell(1, 8).Value) Then
            cell(1, 8).Value = "bk"
            cell(1, 8).Font.ColorIndex = 3
            cell(1, 8).Font.Bold = True
    End If
    End If

    
       If cell.Value <> cell(1, 4).Value And cell(1, 9).Value = "Ponticello inseribile" Then
            cell(1, 9).Value = "Conduttore/filo"
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
            If IsEmpty(cell(1, 8).Value) Then
            cell(1, 8).Value = "bk"
            cell(1, 8).Font.ColorIndex = 3
            cell(1, 8).Font.Bold = True
    End If
    End If
        If cell.Value <> cell(1, 4).Value And cell(1, 9).Value = "Ponticello a filo" Then
            cell(1, 9).Value = "Conduttore/filo"
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
            If IsEmpty(cell(1, 8).Value) Then
            cell(1, 8).Value = "bk"
            cell(1, 8).Font.ColorIndex = 3
            cell(1, 8).Font.Bold = True
    End If
    End If
        
            If cell.Value <> cell(1, 4).Value And cell(1, 9).Value = "Insertable jumper" Then
            cell(1, 9).Value = "Conductor / wire"
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
             If IsEmpty(cell(1, 8).Value) Then
            cell(1, 8).Value = "bk"
            cell(1, 8).Font.ColorIndex = 3
            cell(1, 8).Font.Bold = True
    End If
    End If
      
                 If cell.Value <> cell(1, 4).Value And cell(1, 9).Value = "Saddle jumper" Then
            cell(1, 9).Value = "Conductor / wire"
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
            If IsEmpty(cell(1, 8).Value) Then
            cell(1, 8).Value = "bk"
            cell(1, 8).Font.ColorIndex = 3
            cell(1, 8).Font.Bold = True
    End If
    End If
         
                If cell.Value <> cell(1, 4).Value And cell(1, 9).Value = "Wire jumper" Then
            cell(1, 9).Value = "Conductor / wire"
            cell(1, 9).Font.ColorIndex = 3
            cell(1, 9).Font.Bold = True
            If IsEmpty(cell(1, 8).Value) Then
            cell(1, 8).Value = "bk"
            cell(1, 8).Font.ColorIndex = 3
            cell(1, 8).Font.Bold = True
    End If
    End If

    
 Next
        
        
 '------------------XDA -------------------------

Set myplage = Range("A15:A1000")

    For Each cell In myplage

        
           If Left(cell.Value, 3) = "XDA" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) >= 1 And cell(1, 9).Value = "Ponticello a staffa" Then
        cell(1, 7).ClearContents
        cell(1, 8).ClearContents
        End If
        End If
   If Left(cell.Value, 3) = "XDA" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) >= 1 And cell(1, 9).Value = "Ponticello inseribile" Then
         cell(1, 7).ClearContents
            cell(1, 8).ClearContents

        End If
        End If
        
           If Left(cell.Value, 3) = "XDA" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) >= 1 And cell(1, 9).Value = "Insertable jumper" Then
        cell(1, 7).ClearContents
        cell(1, 8).ClearContents
        End If
        End If
   If Left(cell.Value, 3) = "XDA" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) >= 1 And cell(1, 9).Value = "Saddle jumper" Then
        cell(1, 7).ClearContents
        cell(1, 8).ClearContents
        End If
        End If
        
 Next
        
'------------------XDV -------------------------

Set myplage = Range("A15:A1000")

    For Each cell In myplage

        
           If Left(cell.Value, 3) = "XDV" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) >= 1 And cell(1, 9).Value = "Ponticello a staffa" Then
        cell(1, 7).ClearContents
        cell(1, 8).ClearContents
        End If
        End If
   If Left(cell.Value, 3) = "XDV" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) >= 1 And cell(1, 9).Value = "Ponticello inseribile" Then
         cell(1, 7).ClearContents
        cell(1, 8).ClearContents
        End If
        End If
        
           If Left(cell.Value, 3) = "XDV" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) >= 1 And cell(1, 9).Value = "Insertable jumper" Then
        cell(1, 7).ClearContents
        cell(1, 8).ClearContents
        End If
        End If
   If Left(cell.Value, 3) = "XDV" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) >= 1 And cell(1, 9).Value = "Saddle jumper" Then
        cell(1, 7).ClearContents
        cell(1, 8).ClearContents
        End If
        End If
        
 Next
        
     '------------------XDC -------------------------

    
Set myplage = Range("A15:A1000")

    For Each cell In myplage

        
           If Left(cell.Value, 3) = "XDC" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) >= 1 And cell(1, 9).Value = "Ponticello a staffa" Then
        cell(1, 9).Value = "Ponticello a filo"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
   If Left(cell.Value, 3) = "XDC" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) >= 1 And cell(1, 9).Value = "Ponticello inseribile" Then
        cell(1, 9).Value = "Ponticello a filo"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
        
           If Left(cell.Value, 3) = "XDC" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) >= 1 And cell(1, 9).Value = "Insertable jumper" Then
        cell(1, 9).Value = "Wire jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
   If Left(cell.Value, 3) = "XDC" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) >= 1 And cell(1, 9).Value = "Saddle jumper" Then
        cell(1, 9).Value = "Wire jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
        
                 If IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "XDC" And (cell(1, 9).Value = "Wire jumper" Or cell(1, 9).Value = "Ponticello a filo") Then
                 
        XDC = InputBox("Please add cross-section of conductors between" & cell(1, 3) & " and " & cell(1, 6), "Wire jumper  between " & cell(1, 3) & " and " & cell(1, 6), wire)
        cell(1, 7).Value = XDC
        cell(1, 8).Value = "bk"
        End If

    Next
 
 
 
 
     '------------------XDM -------------------------

Set myplage = Range("A15:A1000")

    For Each cell In myplage

        
           If Left(cell.Value, 3) = "XDM" And cell.Value = cell(1, 4).Value And cell(1, 9).Value = "Ponticello a staffa" Then
        cell(1, 9).Value = "Ponticello a filo"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        
            
           If Left(cell.Value, 3) = "XDM" And cell.Value = cell(1, 4).Value And cell(1, 9).Value = "Ponticello inseribile" Then
        cell(1, 9).Value = "Ponticello a filo"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
            

        
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

Set myplage = Range("A15:A1000")

    For Each cell In myplage

        
           If Left(cell.Value, 2) = "PG" And cell.Value = cell(1, 4).Value And cell(1, 9).Value = "Ponticello a staffa" Then
        cell(1, 9).Value = "Ponticello a filo"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        
            
           If Left(cell.Value, 2) = "PG" And cell.Value = cell(1, 4).Value And cell(1, 9).Value = "Ponticello inseribile" Then
        cell(1, 9).Value = "Ponticello a filo"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
            

        
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

Set myplage = Range("A15:A1000")

    For Each cell In myplage

        
           If Left(cell.Value, 2) = "SF" And cell.Value = cell(1, 4).Value And cell(1, 9).Value = "Ponticello a staffa" Then
        cell(1, 9).Value = "Ponticello a filo"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        
            
           If Left(cell.Value, 2) = "SF" And cell.Value = cell(1, 4).Value And cell(1, 9).Value = "Ponticello inseribile" Then
        cell(1, 9).Value = "Ponticello a filo"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
            

        
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

Set myplage = Range("A15:A1000")

    For Each cell In myplage

        
           If Left(cell.Value, 2) = "BT" And cell.Value = cell(1, 4).Value And cell(1, 9).Value = "Ponticello a staffa" Then
        cell(1, 9).Value = "Ponticello a filo"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        
            
           If Left(cell.Value, 2) = "BT" And cell.Value = cell(1, 4).Value And cell(1, 9).Value = "Ponticello inseribile" Then
        cell(1, 9).Value = "Ponticello a filo"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
            

        
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

Set myplage = Range("A15:A1000")

    For Each cell In myplage

        
           If Left(cell.Value, 2) = "TB" And cell.Value = cell(1, 4).Value And cell(1, 9).Value = "Ponticello a staffa" Then
        cell(1, 9).Value = "Ponticello a filo"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        
            
           If Left(cell.Value, 2) = "TB" And cell.Value = cell(1, 4).Value And cell(1, 9).Value = "Ponticello inseribile" Then
        cell(1, 9).Value = "Ponticello a filo"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
            

        
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

Set myplage = Range("A15:A1000")

    For Each cell In myplage


            If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "XDI" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) = 1 And cell(1, 9).Value = "Conductor / wire" Then
        cell(1, 7).ClearContents
        cell(1, 8).ClearContents
        cell(1, 9).Value = "Saddle jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
                  If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "XDI" And cell.Value = cell(1, 4).Value Then
                  If Abs(cell(1, 2).Value - cell(1, 5).Value) = 1 And cell(1, 9).Value = "Wire jumper" Then
        cell(1, 7).ClearContents
        cell(1, 8).ClearContents
         cell(1, 9).Value = "Saddle jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
        
        
        
        
        If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "XDI" And cell.Value = cell(1, 4).Value Then
        If Abs(cell(1, 2).Value - cell(1, 5).Value) = 1 And cell(1, 9).Value = "Conduttore/filo" Then
         cell(1, 7).ClearContents
        cell(1, 8).ClearContents
        cell(1, 9).Value = "Ponticello a staffa"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True

        End If
        End If
               
               If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "XDI" And cell.Value = cell(1, 4).Value Then
        If Abs(cell(1, 2).Value - cell(1, 5).Value) = 1 And cell(1, 9).Value = "Ponticello a filo" Then
        cell(1, 7).ClearContents
        cell(1, 8).ClearContents
        cell(1, 9).Value = "Ponticello a staffa"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        
        End If
        End If
        
           If Left(cell.Value, 3) = "XDI" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) > 1 And cell(1, 9).Value = "Ponticello a staffa" Then
        cell(1, 9).Value = "Ponticello a filo"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
        
        If Left(cell.Value, 3) = "XDI" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) > 1 And cell(1, 9).Value = "Ponticello inseribile" Then
        cell(1, 9).Value = "Ponticello a filo"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
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
         If Abs(cell(1, 2).Value - cell(1, 5).Value) = 1 And cell(1, 9).Value = "Ponticello a staffa" Then
        cell(1, 7).ClearContents
        cell(1, 8).ClearContents
        End If
        End If
                  If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "XDI" And cell.Value = cell(1, 4).Value Then
                  If Abs(cell(1, 2).Value - cell(1, 5).Value) = 1 And cell(1, 9).Value = "Ponticello inseribile" Then
        cell(1, 7).ClearContents
        cell(1, 8).ClearContents

        End If
        End If
        
        
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
        XDI = InputBox("Please add cross-section of conductors between" & cell(1, 3) & " and " & cell(1, 6), "Wire jumper  between " & cell(1, 3) & " and " & cell(1, 6), wire)
        cell(1, 7).Value = XDI
        cell(1, 8).Value = "bk"
        End If
        
        
        Next
    '--------------------------------------RAR------------------------------
    Set myplage = Range("A15:A1000")

    For Each cell In myplage

          If Left(cell.Value, 3) = "RAR" Then
          answer = MsgBox("Do you want to clear section between " & cell(1, 3).Value & " and " & cell(1, 6).Value, vbYesNo + vbQuestion, "RAR section")
          If answer = vbYes Then
            cell(1, 7).ClearContents
            cell(1, 8).ClearContents
         End If
            End If
 Next
                 
                 
 Set myplage = Range("D15:D1000")

    For Each cell In myplage
    

          If Left(cell.Value, 3) = "RAR" Then
          answer = MsgBox("Do you want to clear section between " & cell(1, 3).Value & " and " & cell.Offset(1, -1).Value, vbYesNo + vbQuestion, "RAR section")
        If answer = vbYes Then
            cell(1, 4).ClearContents
            cell(1, 5).ClearContents
        End If
        End If
Next

'------------------XDX -------------------------

Set myplage = Range("A15:A1000")

    For Each cell In myplage


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
        If Abs(cell(1, 2).Value - cell(1, 5).Value) = 1 And cell(1, 9).Value = "Conduttore/filo" Then
        cell(1, 9).Value = "Ponticello a staffa"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        cell(1, 7).ClearContents
        cell(1, 8).ClearContents
        End If
         End If
               
               If Left(cell.Value, 3) = "XDX" And cell.Value = cell(1, 4).Value Then
        If Abs(cell(1, 2).Value - cell(1, 5).Value) = 1 And cell(1, 9).Value = "Ponticello a filo" Then
        cell(1, 9).Value = "Ponticello a staffa"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        cell(1, 7).ClearContents
        cell(1, 8).ClearContents
        End If
        End If
        
           If Left(cell.Value, 3) = "XDX" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) > 1 And cell(1, 9).Value = "Ponticello a staffa" Then
        cell(1, 9).Value = "Ponticello a filo"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
   If Left(cell.Value, 3) = "XDX" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) > 1 And cell(1, 9).Value = "Ponticello inseribile" Then
        cell(1, 9).Value = "Ponticello a filo"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
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
         If Abs(cell(1, 2).Value - cell(1, 5).Value) = 1 And cell(1, 9).Value = "Ponticello a staffa" Then
        cell(1, 7).ClearContents
        cell(1, 8).ClearContents
        End If
        End If
                  If Not IsEmpty(cell(1, 7).Value) And Left(cell.Value, 3) = "XDX" And cell.Value = cell(1, 4).Value Then
                  If Abs(cell(1, 2).Value - cell(1, 5).Value) = 1 And cell(1, 9).Value = "Ponticello inseribile" Then
        cell(1, 7).ClearContents
        cell(1, 8).ClearContents
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
        XDX = InputBox("Please add cross-section of conductors between" & cell(1, 3) & " and " & cell(1, 6), "Wire jumper between " & cell(1, 3) & " and " & cell(1, 6), wire)
        cell(1, 7).Value = XDX
        cell(1, 8).Value = "bk"
        End If
        
        
 
 
  '---------------------------Wire Bridges----------------------------------------------
  
  

   If cell.Value = "XDB1" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) >= 1 And cell(1, 9).Value = "Ponticello a staffa" Then
        cell(1, 9).Value = "Ponticello a filo"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
   If cell.Value = "XDB1" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) >= 1 And cell(1, 9).Value = "Ponticello inseribile" Then
        cell(1, 9).Value = "Ponticello a filo"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
        
           If cell.Value = "XDB1" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) >= 1 And cell(1, 9).Value = "Insertable jumper" Then
        cell(1, 9).Value = "Wire jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
   If cell.Value = "XDB1" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) >= 1 And cell(1, 9).Value = "Saddle jumper" Then
        cell(1, 9).Value = "Wire jumper"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
        
          If IsEmpty(cell(1, 7).Value) And cell.Value = "XDB1" And (cell(1, 9).Value = "Wire jumper" Or cell(1, 9).Value = "Ponticello a filo") Then
        XDB1 = InputBox("Please add cross-section of conductors between" & cell(1, 3) & " and " & cell(1, 6), "Wire jumper between " & cell(1, 3) & " and " & cell(1, 6), wire)
        cell(1, 7).Value = XDB1
        cell(1, 8).Value = "bk"
        cell(1, 7).Font.ColorIndex = 3
        cell(1, 7).Font.Bold = True
        End If
        
        
           If cell.Value = "XDE" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) >= 1 And cell(1, 9).Value = "Ponticello a staffa" Then
        cell(1, 9).Value = "Ponticello a filo"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
   If cell.Value = "XDE" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) >= 1 And cell(1, 9).Value = "Ponticello inseribile" Then
        cell(1, 9).Value = "Ponticello a filo"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
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
        XDE = InputBox("Please add cross-section of conductors between" & cell(1, 3) & " and " & cell(1, 6), "Wire jumper between " & cell(1, 3) & " and " & cell(1, 6), wire)
        cell(1, 7).Value = XDE
        cell(1, 8).Value = "bk"
        End If
              
        
        
           If cell.Value = "XDT" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) >= 1 And cell(1, 9).Value = "Ponticello a staffa" Then
        cell(1, 9).Value = "Ponticello a filo"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
        End If
   If cell.Value = "XDT" And cell.Value = cell(1, 4).Value Then
         If Abs(cell(1, 2).Value - cell(1, 5).Value) >= 1 And cell(1, 9).Value = "Ponticello inseribile" Then
        cell(1, 9).Value = "Ponticello a filo"
        cell(1, 9).Font.ColorIndex = 3
        cell(1, 9).Font.Bold = True
        End If
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
        XDT = InputBox("Please add cross-section of conductors between" & cell(1, 3) & " and " & cell(1, 6), "Wire jumper between " & cell(1, 3) & " and " & cell(1, 6), wire)
        cell(1, 7).Value = XDT
        cell(1, 8).Value = "bk"
        End If
        
  Next
 
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True


        
 End Sub
  
   


   
