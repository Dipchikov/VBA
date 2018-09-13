Attribute VB_Name = "Translate"
Sub translate()
Dim lr As Long
lr = Range("A" & Rows.Count).End(xlUp).Row
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

 '---------Italian to English----------------
Set MyPlage = Range("I15:I" & lr)

    For Each cell In MyPlage
    
        If cell.Value = "Collegamento diretto" Then
        cell.Value = "Direct connection"
        End If
        If cell.Value = "Interno" Then
        cell.Value = "Internal"
        End If
        
        If cell.Value = "Ponticello a staffa" Then
        cell.Value = "Saddle jumper"
        End If
        
        
        If cell.Value = "Ponticello a filo" Then
        cell.Value = "Wire jumper"
        End If
                
  
        If cell.Value = "Ponticello inseribile" Then
        cell.Value = "Insertable jumper"
        End If
        If cell.Value = "Conduttore/filo" Then
        cell.Value = "Conductor / wire"
        End If
         If cell.Value = "Conduttore/filo" Then
        cell.Value = "Conductor / wire"
        End If
        
        
        If cell.Value = "Conduttore / filo" Then
        cell.Value = "Conductor / wire"
        End If
        
        '-----------Spanish translation--------------
                If cell.Value = "Conductor/cable" Then
        cell.Value = "Conductor / wire"
        End If
        
                If cell.Value = "Puente de regleta" Then
        cell.Value = "Saddle jumper"
        End If
        
              If cell.Value = "Puente de hilo" Then
        cell.Value = "Wire jumper"
        End If

        Next
        
        '---------fix cable----------------
        
    Set MyPlage = Range("H15:H" & lr)
    For Each cell In MyPlage
  
        
    If (cell.Value = "black" Or cell.Value = "BLACK") Then
        cell.Value = "bk"
        End If
        
    Next
        
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
End Sub
