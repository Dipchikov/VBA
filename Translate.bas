Attribute VB_Name = "Translate"
Sub translate()

Set MyPlage = Range("I15:I1000")

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
        
        
        Next
End Sub
