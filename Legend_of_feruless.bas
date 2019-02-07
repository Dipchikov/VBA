Attribute VB_Name = "Legend_of_feruless"
Sub Legend_of_feruless()

Dim lr As Long
Dim ferulle_10 As String
Dim ferulle_15 As String
Dim ferulle_25 As String
Dim ferulle_40 As String

Set Data = Sheets("Type of cables")
Set ref = Sheets("BOM")

If Komax.ferulles_Yes.Value = True Then

 If Ferrules.CheckBox1.Value = True Then
 ferulle_10 = Ferrules.Ferrules_10.Value
 End If

 If Ferrules.CheckBox2.Value = True Then
 ferulle_15 = Ferrules.Ferrules_15.Value
 End If

 If Ferrules.CheckBox3.Value = True Then
 ferulle_25 = Ferrules.Ferrules_25.Value
 End If
 
 If Ferrules.CheckBox4.Value = True Then
 ferulle_40 = Ferrules.Ferrules_40.Value
 End If
 
 Else
 
 ferulle_10 = ""
 ferulle_15 = ""
 ferulle_25 = ""
 ferulle_40 = ""
 
End If

Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False


Application.ScreenUpdating = False
lr = Range("A" & Rows.Count).End(xlUp).Row



On Error Resume Next
'Columns("T").EntireColumn.Delete
Range("T15:U1048576").Select
Selection.ClearContents

'------------------Inside Wiring -------------------------


Set MyPlage = Range("A15:A" & lr)

    For Each cell In MyPlage
    
If Komax.XDC.Value = True Then
        If Left(cell.Value, 3) = "XDC" Then
        cell(1, 20).Value = 10
        End If
        If Left(cell.Value, 3) = "XDC" And cell(1, 7).Value = 1 Then
        cell(1, 20).Value = 11
        End If
        If Left(cell.Value, 3) = "XDC" And cell(1, 7).Value = 1.5 Then
        cell(1, 20).Value = 12
        End If

Else
        If Left(cell.Value, 3) = "XDC" Then
        cell(1, 20).Value = ""
        End If
End If
          If Left(cell.Value, 2) = "BT" And Not cell(1, 7).Value = 1 Then
        cell(1, 20).Value = 10
        End If
           If Left(cell.Value, 2) = "BT" And cell(1, 7).Value = 1 Then
        cell(1, 20).Value = 11
        End If
           If Left(cell.Value, 2) = "PE" Then
        cell(1, 20).Value = 10
        End If
        
        If Left(cell.Value, 2) = "PE" And cell(1, 7).Value = 1 Then
        cell(1, 20).Value = 11
        End If
        
                  If Left(cell.Value, 2) = "IE" Then
        cell(1, 20).Value = 10
        End If
        
                If Left(cell.Value, 2) = "IE" And cell(1, 7).Value = 1 Then
        cell(1, 20).Value = 11
        End If
        
                  If Left(cell.Value, 2) = "EA" Then
        cell(1, 20).Value = 10
        End If
                If Left(cell.Value, 2) = "EA" And cell(1, 7).Value = 1 Then
        cell(1, 20).Value = 11
        End If
        
            If Left(cell.Value, 2) = "BR" Then
        cell(1, 20).Value = 10
        End If
                If Left(cell.Value, 2) = "BR" And cell(1, 7).Value = 1 Then
        cell(1, 20).Value = 11
        End If

          If Left(cell.Value, 2) = "BM" Then
        cell(1, 20).Value = 10
        End If
                If Left(cell.Value, 2) = "BM" And cell(1, 7).Value = 1 Then
        cell(1, 20).Value = 11
        End If
        
                  If Left(cell.Value, 2) = "BX" Then
        cell(1, 20).Value = 10
        End If
                If Left(cell.Value, 2) = "BX" And cell(1, 7).Value = 1 Then
        cell(1, 20).Value = 11
        End If
        
                          If Left(cell.Value, 2) = "TS" Then
        cell(1, 20).Value = 10
        End If
                If Left(cell.Value, 2) = "TS" And cell(1, 7).Value = 1 Then
        cell(1, 20).Value = 11
        End If
        
           If cell.Value = "AA1" Then
        cell(1, 20).Value = 10
        End If
                If cell.Value = "AA1" And cell(1, 7).Value = 1 Then
        cell(1, 20).Value = 11
        End If
        
                   If cell.Value = "AA2" Then
        cell(1, 20).Value = 10
        End If
        
             If cell.Value = "AA2" And cell(1, 7).Value = 1 Then
        cell(1, 20).Value = 11
        End If
        
                  If cell.Value = "AA3" Then
        cell(1, 20).Value = 10
        End If
              If cell.Value = "AA3" And cell(1, 7).Value = 1 Then
        cell(1, 20).Value = 11
        End If
                
                   If cell.Value = "AA4" Then
        cell(1, 20).Value = 10
        End If
             If cell.Value = "AA4" And cell(1, 7).Value = 1 Then
        cell(1, 20).Value = 11
        End If
        
                If cell.Value = "XDB1" Then
        cell(1, 20).Value = ""
        End If
                If Left(cell.Value, 3) = "XDE" Then
        cell(1, 20).Value = ""
        End If
        
    
                If Left(cell.Value, 3) = "XDT" Then
        cell(1, 20).Value = ""
        End If
                If Left(cell.Value, 3) = "PFV" Then
        cell(1, 20).Value = 10
        End If
              If Left(cell.Value, 3) = "PFV" And cell(1, 7).Value = 1 Then
        cell(1, 20).Value = 11
        End If
        
            If Left(cell.Value, 3) = "RAD" Then
        cell(1, 20).Value = 10
        End If
           If Left(cell.Value, 3) = "RAD" And cell(1, 7).Value = 1 Then
        cell(1, 20).Value = 11
        End If
           If Left(cell.Value, 3) = "RAD" And cell(1, 7).Value = 1.5 Then
        cell(1, 20).Value = 12
        End If
        
        If Left(cell.Value, 3) = "FCM" And Not (cell(1, 2).Value = 13 Or cell(1, 2).Value = 14 Or cell(1, 2).Value = 21 Or cell(1, 2).Value = 22 Or cell(1, 2).Value = 96 Or cell(1, 2).Value = 95 Or cell(1, 2).Value = 98) Then
            cell(1, 20).Value = 14
        End If
                If Left(cell.Value, 3) = "FCM" And (cell(1, 2).Value = 13 Or cell(1, 2).Value = 14 Or cell(1, 2).Value = 21 Or cell(1, 2).Value = 22 Or cell(1, 2).Value = 96 Or cell(1, 2).Value = 95 Or cell(1, 2).Value = 98) Then
            cell(1, 20).Value = 11
        End If
                
        If Left(cell.Value, 2) = "TB" Then
        cell(1, 20).Value = 10
        End If
            If Left(cell.Value, 2) = "TB" And cell(1, 7).Value = 1 Then
        cell(1, 20).Value = 11
        End If
        
If Komax.XDX.Value = True Then
        If Left(cell.Value, 3) = "XDX" Then
        cell(1, 20).Value = 10
        End If
        If Left(cell.Value, 3) = "XDX" And cell(1, 7).Value = 1 Then
        cell(1, 20).Value = 11
        End If
        If Left(cell.Value, 3) = "XDX" And cell(1, 7).Value = 1.5 Then
        cell(1, 20).Value = 12
        End If

Else
        If Left(cell.Value, 3) = "XDX" Then
        cell(1, 20).Value = ""
        End If
End If

If Komax.PHOENIX.Value = True Then
        If cell.Value = "XDA" Then
        cell(1, 20).Value = 12
        End If
        Else
        If cell.Value = "XDA" Then
        cell(1, 20).Value = 14
        End If
 End If
        
    If Komax.PHOENIX.Value = True Then
        If cell.Value = "XDV" Then
        cell(1, 20).Value = 10
        End If
        Else
        If cell.Value = "XDV" Then
        cell(1, 20).Value = 14
        End If
End If
If Komax.XDI.Value = True Then
        If Left(cell.Value, 3) = "XDI" Then
        cell(1, 20).Value = 10
        End If
        If Left(cell.Value, 3) = "XDI" And cell(1, 7).Value = 1 Then
        cell(1, 20).Value = 11
        End If
        If Left(cell.Value, 3) = "XDI" And cell(1, 7).Value = 1.5 Then
        cell(1, 20).Value = 12
        End If

Else
        If Left(cell.Value, 3) = "XDI" Then
        cell(1, 20).Value = ""
        End If
End If


        If cell.Value = "K1" Then
        cell(1, 20).Value = 10
        End If
        If cell.Value = "K1" And cell(1, 7).Value = 1 Then
        cell(1, 20).Value = 11
        End If
        
        
        If cell.Value = "K2" Then
        cell(1, 20).Value = 10
        End If
                If cell.Value = "K2" And cell(1, 7).Value = 1 Then
        cell(1, 20).Value = 11
        End If
        
        If cell.Value = "K3" Then
        cell(1, 20).Value = 10
        End If
               If cell.Value = "K3" And cell(1, 7).Value = 1 Then
        cell(1, 20).Value = 11
        End If
        
        If cell.Value = "K4" Then
        cell(1, 20).Value = 10
        End If
        If cell.Value = "K4" And cell(1, 7).Value = 1 Then
        cell(1, 20).Value = 11
        End If
        
        If Left(cell.Value, 2) = "KA" Then
        cell(1, 20).Value = 10
        End If
        If Left(cell.Value, 2) = "KA" And cell(1, 7).Value = 1 Then
        cell(1, 20).Value = 11
        End If
                    
             If Left(cell.Value, 2) = "KF" Then
        cell(1, 20).Value = 10
        End If
        If Left(cell.Value, 2) = "KF" And cell(1, 7).Value = 1 Then
        cell(1, 20).Value = 11
        End If
        
            If Left(cell.Value, 3) = "RAA" Then
        cell(1, 20).Value = 10
        End If
        If Left(cell.Value, 3) = "RAA" And cell(1, 7).Value = 1 Then
        cell(1, 20).Value = 11
        End If



        If Left(cell.Value, 3) = "TFS" Then
        cell(1, 20).Value = 10
        End If
        If Left(cell.Value, 3) = "TFS" And cell(1, 7).Value = 1 Then
        cell(1, 20).Value = 11
        End If
        
        If Left(cell.Value, 3) = "TFM" Then
        cell(1, 20).Value = 10
        End If
        If Left(cell.Value, 3) = "TFM" And cell(1, 7).Value = 1 Then
        cell(1, 20).Value = 11
        End If
        

        
        
If Komax.RAR.Value = True Then
        If Left(cell.Value, 3) = "RAR" Then
        cell(1, 20).Value = 10
        End If
        If Left(cell.Value, 3) = "RAR" And cell(1, 7).Value = 1 Then
        cell(1, 20).Value = 11
        End If
        If Left(cell.Value, 3) = "RAR" And cell(1, 7).Value = 1.5 Then
        cell(1, 20).Value = 12
        End If

Else
        If Left(cell.Value, 3) = "RAR" Then
        cell(1, 20).Value = ""
        End If
End If
                   
        If Left(cell.Value, 2) = "XE" Then
        cell(1, 20).Value = 10
        End If
        
           If Left(cell.Value, 3) = "XDS" Then
        cell(1, 20).Value = 10
        End If
          If Left(cell.Value, 3) = "XDS" And cell(1, 7).Value = 1 Then
        cell(1, 20).Value = 11
        End If
        
    Next
    

    '----------------------------Door Wireing ----------------------------
    
    
    Set MyPlage = Range("A15:A" & lr)
        For Each cell In MyPlage
        
            If Left(cell.Value, 3) = "SPM" Then
        cell(1, 20).Value = 10
        End If
        If Left(cell.Value, 3) = "SPM" And cell(1, 7).Value = 1 Then
        cell(1, 20).Value = 11
        End If
           If Left(cell.Value, 3) = "STF" Then
        cell(1, 20).Value = 10
        End If
         If Left(cell.Value, 3) = "STF" And cell(1, 7).Value = 1 Then
        cell(1, 20).Value = 11
        End If
          
        
          
        If Left(cell.Value, 3) = "XDM" Then
        cell(1, 20).Value = 10
        End If
        
                

          If (Left(cell.Value, 2) = "PG" Or Left(cell.Value, 2) = "PF") Then
        cell(1, 20).Value = 10
        End If
                If (Left(cell.Value, 2) = "PG" Or Left(cell.Value, 2) = "PF") And cell(1, 7).Value = 1 Then
        cell(1, 20).Value = 11
        End If
        
            If Left(cell.Value, 2) = "SF" Then
        cell(1, 20).Value = 10
        End If
            If Left(cell.Value, 2) = "SF" And cell(1, 7).Value = 1 Then
        cell(1, 20).Value = 11
        End If
        
        '---------------LOCKOUT RELAY---------------------
        
        
                If Left(cell.Value, 3) = "K86" Then
        cell(1, 20).Value = 10
        End If
                    If Left(cell.Value, 3) = "K86" And cell(1, 7).Value = 1 Then
        cell(1, 20).Value = 11
        End If


    Next

 '---------------------Wireing - 'Ref protection-----------------
 
    Set MyPlage = Range("A15:A" & lr)

    For Each cell In MyPlage

        If Left(cell.Value, 2) = "AA" And (Left(cell(1, 2).Value, 5) = "-X130" Or Left(cell(1, 2).Value, 5) = "-X327" Or Left(cell(1, 2).Value, 5) = "-X329" Or Left(cell(1, 2).Value, 5) = "-X321" Or Left(cell(1, 2).Value, 5) = "-X324" Or Left(cell(1, 2).Value, 5) = "-X316" Or Left(cell(1, 2).Value, 5) = "-X319" Or Left(cell(1, 2).Value, 5) = "-X304" Or Left(cell(1, 2).Value, 5) = "-X307" Or Left(cell(1, 2).Value, 5) = "-X309" Or Left(cell(1, 2).Value, 5) = "-X410" Or Left(cell(1, 2).Value, 5) = "-X304" Or Left(cell(1, 2).Value, 5) = "-X326" Or Left(cell(1, 2).Value, 5) = "-X331" Or Left(cell(1, 2).Value, 5) = "-X336" Or Left(cell(1, 2).Value, 5) = "-X334" Or Left(cell(1, 2).Value, 5) = "-X339") Then
            cell(1, 20).Value = 14
        End If
        
            If Left(cell.Value, 2) = "AA" And Not (Left(cell(1, 2).Value, 5) = "-X130" Or Left(cell(1, 2).Value, 5) = "-X327" Or Left(cell(1, 2).Value, 5) = "-X329" Or Left(cell(1, 2).Value, 5) = "-X321" Or Left(cell(1, 2).Value, 5) = "-X324" Or Left(cell(1, 2).Value, 5) = "-X316" Or Left(cell(1, 2).Value, 5) = "-X319" Or Left(cell(1, 2).Value, 5) = "-X304" Or Left(cell(1, 2).Value, 5) = "-X307" Or Left(cell(1, 2).Value, 5) = "-X309" Or Left(cell(1, 2).Value, 5) = "-X410" Or Left(cell(1, 2).Value, 5) = "-X304" Or Left(cell(1, 2).Value, 5) = "-X326" Or Left(cell(1, 2).Value, 5) = "-X331" Or Left(cell(1, 2).Value, 5) = "-X336" Or Left(cell(1, 2).Value, 5) = "-X334" Or Left(cell(1, 2).Value, 5) = "-X339") And cell(1, 7).Value = 4 Then
            cell(1, 20).Value = 12
        End If
                If Left(cell.Value, 2) = "AA" And Not (Left(cell(1, 2).Value, 5) = "-X130" Or Left(cell(1, 2).Value, 5) = "-X327" Or Left(cell(1, 2).Value, 5) = "-X329" Or Left(cell(1, 2).Value, 5) = "-X321" Or Left(cell(1, 2).Value, 5) = "-X324" Or Left(cell(1, 2).Value, 5) = "-X316" Or Left(cell(1, 2).Value, 5) = "-X319" Or Left(cell(1, 2).Value, 5) = "-X304" Or Left(cell(1, 2).Value, 5) = "-X307" Or Left(cell(1, 2).Value, 5) = "-X309" Or Left(cell(1, 2).Value, 5) = "-X410" Or Left(cell(1, 2).Value, 5) = "-X304" Or Left(cell(1, 2).Value, 5) = "-X326" Or Left(cell(1, 2).Value, 5) = "-X331" Or Left(cell(1, 2).Value, 5) = "-X336" Or Left(cell(1, 2).Value, 5) = "-X334" Or Left(cell(1, 2).Value, 5) = "-X339") And cell(1, 7).Value = 2.5 Then
            cell(1, 20).Value = 10
        End If
        
            If Left(cell.Value, 2) = "AA" And Not (Left(cell(1, 2).Value, 5) = "-X130" Or Left(cell(1, 2).Value, 5) = "-X327" Or Left(cell(1, 2).Value, 5) = "-X329" Or Left(cell(1, 2).Value, 5) = "-X321" Or Left(cell(1, 2).Value, 5) = "-X324" Or Left(cell(1, 2).Value, 5) = "-X316" Or Left(cell(1, 2).Value, 5) = "-X319" Or Left(cell(1, 2).Value, 5) = "-X304" Or Left(cell(1, 2).Value, 5) = "-X307" Or Left(cell(1, 2).Value, 5) = "-X309" Or Left(cell(1, 2).Value, 5) = "-X410" Or Left(cell(1, 2).Value, 5) = "-X304" Or Left(cell(1, 2).Value, 5) = "-X326" Or Left(cell(1, 2).Value, 5) = "-X331" Or Left(cell(1, 2).Value, 5) = "-X336" Or Left(cell(1, 2).Value, 5) = "-X334" Or Left(cell(1, 2).Value, 5) = "-X339") And cell(1, 7).Value = 1.5 Then
            cell(1, 20).Value = 12
        End If
            If Left(cell.Value, 2) = "AA" And Not (Left(cell(1, 2).Value, 5) = "-X130" Or Left(cell(1, 2).Value, 5) = "-X327" Or Left(cell(1, 2).Value, 5) = "-X329" Or Left(cell(1, 2).Value, 5) = "-X321" Or Left(cell(1, 2).Value, 5) = "-X324" Or Left(cell(1, 2).Value, 5) = "-X316" Or Left(cell(1, 2).Value, 5) = "-X319" Or Left(cell(1, 2).Value, 5) = "-X304" Or Left(cell(1, 2).Value, 5) = "-X307" Or Left(cell(1, 2).Value, 5) = "-X309" Or Left(cell(1, 2).Value, 5) = "-X410" Or Left(cell(1, 2).Value, 5) = "-X304" Or Left(cell(1, 2).Value, 5) = "-X326" Or Left(cell(1, 2).Value, 5) = "-X331" Or Left(cell(1, 2).Value, 5) = "-X336" Or Left(cell(1, 2).Value, 5) = "-X334" Or Left(cell(1, 2).Value, 5) = "-X339") And cell(1, 7).Value = 1 Then
            cell(1, 20).Value = 11
        End If
        
        
        
        If (Left(cell.Value, 2) = "BC" Or Left(cell.Value, 2) = "BE") And Not cell(1, 7).Value = 4 Then
            cell(1, 20).Value = 10
        End If
            If (Left(cell.Value, 2) = "BC" Or Left(cell.Value, 2) = "BE") And cell(1, 7).Value = 4 Then
            cell(1, 20).Value = 12
        End If
           If (Left(cell.Value, 2) = "BC" Or Left(cell.Value, 2) = "BE") And cell(1, 7).Value = 1 Then
        cell(1, 20).Value = 11
        End If
        


         
            If Left(cell.Value, 3) = "BAR" And Not cell(1, 7).Value = 4 Then
            cell(1, 20).Value = 10
        End If
            If Left(cell.Value, 3) = "BAR" And cell(1, 7).Value = 4 Then
            cell(1, 20).Value = 12
        End If
               If Left(cell.Value, 3) = "BAR" And cell(1, 7).Value = 1 Then
        cell(1, 20).Value = 11
        End If
        
        
        
        
  '------------------Ferrules -------------------------




If Komax.ferulles_Yes.Value = True Then

    
    
If Komax.XDC.Value = True Then
 
        If Left(cell.Value, 3) = "XDC" And cell(1, 7).Value = 1 Then
        cell(1, 21).Value = ferulle_10
        End If
        If Left(cell.Value, 3) = "XDC" And cell(1, 7).Value = 1.5 Then
        cell(1, 21).Value = ferulle_15
        End If
        If Left(cell.Value, 3) = "XDC" And cell(1, 7).Value = 2.5 Then
        cell(1, 21).Value = ferulle_25
        End If
        If Left(cell.Value, 3) = "XDC" And cell(1, 7).Value = 4 Then
        cell(1, 21).Value = ferulle_40
        End If

Else
        If Left(cell.Value, 3) = "XDC" Then
        cell(1, 21).Value = ""
        End If
End If


If Komax.XDX.Value = True Then

        If Left(cell.Value, 3) = "XDX" And cell(1, 7).Value = 1 Then
        cell(1, 21).Value = ferulle_10
        End If
        If Left(cell.Value, 3) = "XDX" And cell(1, 7).Value = 1.5 Then
        cell(1, 21).Value = ferulle_15
        End If
        If Left(cell.Value, 3) = "XDX" And cell(1, 7).Value = 2.5 Then
        cell(1, 21).Value = ferulle_25
        End If
        If Left(cell.Value, 3) = "XDX" And cell(1, 7).Value = 4 Then
        cell(1, 21).Value = ferulle_40
        End If

Else
        If Left(cell.Value, 3) = "XDX" Then
        cell(1, 21).Value = ""
        End If
End If


If Komax.XDI.Value = True Then
        If Left(cell.Value, 3) = "XDI" And cell(1, 7).Value = 1 Then
        cell(1, 21).Value = ferulle_10
        End If
        If Left(cell.Value, 3) = "XDI" And cell(1, 7).Value = 1.5 Then
        cell(1, 21).Value = ferulle_15
        End If
        If Left(cell.Value, 3) = "XDI" And cell(1, 7).Value = 2.5 Then
        cell(1, 21).Value = ferulle_25
        End If
        If Left(cell.Value, 3) = "XDI" And cell(1, 7).Value = 4 Then
        cell(1, 21).Value = ferulle_40
        End If

Else
        If Left(cell.Value, 3) = "XDI" Then
        cell(1, 21).Value = ""
        End If
End If

If Komax.RAR.Value = True Then
        If Left(cell.Value, 3) = "RAR" And cell(1, 7).Value = 1 Then
        cell(1, 21).Value = ferulle_10
        End If
        If Left(cell.Value, 3) = "RAR" And cell(1, 7).Value = 1.5 Then
        cell(1, 21).Value = ferulle_15
        End If
        If Left(cell.Value, 3) = "RAR" And cell(1, 7).Value = 2.5 Then
        cell(1, 21).Value = ferulle_25
        End If
        If Left(cell.Value, 3) = "RAR" And cell(1, 7).Value = 4 Then
        cell(1, 21).Value = ferulle_40
        End If

Else
        If Left(cell.Value, 3) = "RAR" Then
        cell(1, 21).Value = ""
        End If
End If

          If cell.Value = "XDB1" Then
        cell(1, 21).Value = ""
        End If
                If Left(cell.Value, 3) = "XDE" Then
        cell(1, 21).Value = ""
        End If
        
                If Left(cell.Value, 3) = "XDT" Then
        cell(1, 21).Value = ""
        End If



        If (Left(cell.Value, 2) = "BT" Or Left(cell.Value, 2) = "PE" Or Left(cell.Value, 2) = "IE" Or Left(cell.Value, 2) = "EA" Or Left(cell.Value, 2) = "XE") And cell(1, 7).Value = 1 Then
        cell(1, 21).Value = ferulle_10
        End If
        If (Left(cell.Value, 2) = "BT" Or Left(cell.Value, 2) = "PE" Or Left(cell.Value, 2) = "IE" Or Left(cell.Value, 2) = "EA" Or Left(cell.Value, 2) = "XE") And cell(1, 7).Value = 1.5 Then
        cell(1, 21).Value = ferulle_15
        End If
        If (Left(cell.Value, 2) = "BT" Or Left(cell.Value, 2) = "PE" Or Left(cell.Value, 2) = "IE" Or Left(cell.Value, 2) = "EA" Or Left(cell.Value, 2) = "XE") And cell(1, 7).Value = 2.5 Then
        cell(1, 21).Value = ferulle_25
        End If
        If (Left(cell.Value, 2) = "BT" Or Left(cell.Value, 2) = "PE" Or Left(cell.Value, 2) = "IE" Or Left(cell.Value, 2) = "EA" Or Left(cell.Value, 2) = "XE") And cell(1, 7).Value = 4 Then
        cell(1, 21).Value = ferulle_40
        End If
        
        
            If (Left(cell.Value, 2) = "BR" Or Left(cell.Value, 2) = "BM" Or Left(cell.Value, 2) = "BX" Or Left(cell.Value, 2) = "TS") And cell(1, 7).Value = 1 Then
        cell(1, 21).Value = ferulle_10
        End If
        If (Left(cell.Value, 2) = "BR" Or Left(cell.Value, 2) = "BM" Or Left(cell.Value, 2) = "BX" Or Left(cell.Value, 2) = "TS") And cell(1, 7).Value = 1.5 Then
        cell(1, 21).Value = ferulle_15
        End If
        If (Left(cell.Value, 2) = "BR" Or Left(cell.Value, 2) = "BM" Or Left(cell.Value, 2) = "BX" Or Left(cell.Value, 2) = "TS") And cell(1, 7).Value = 2.5 Then
        cell(1, 21).Value = ferulle_25
        End If
        If (Left(cell.Value, 2) = "BR" Or Left(cell.Value, 2) = "BM" Or Left(cell.Value, 2) = "BX" Or Left(cell.Value, 2) = "TS") And cell(1, 7).Value = 4 Then
        cell(1, 21).Value = ferulle_40
        End If
        
        
            If (cell.Value = "AA1" Or cell.Value = "AA2" Or cell.Value = "AA3" Or cell.Value = "AA4") And cell(1, 7).Value = 1 Then
        cell(1, 21).Value = ferulle_10
        End If
        If (cell.Value = "AA1" Or cell.Value = "AA2" Or cell.Value = "AA3" Or cell.Value = "AA4") And cell(1, 7).Value = 1.5 Then
        cell(1, 21).Value = ferulle_15
        End If
        If (cell.Value = "AA1" Or cell.Value = "AA2" Or cell.Value = "AA3" Or cell.Value = "AA4") And cell(1, 7).Value = 2.5 Then
        cell(1, 21).Value = ferulle_25
        End If
        If (cell.Value = "AA1" Or cell.Value = "AA2" Or cell.Value = "AA3" Or cell.Value = "AA4") And cell(1, 7).Value = 4 Then
        cell(1, 21).Value = ferulle_40
        End If
        


          If (Left(cell.Value, 3) = "RAD" Or Left(cell.Value, 2) = "TB") And cell(1, 7).Value = 1 Then
        cell(1, 21).Value = ferulle_10
        End If
        If (Left(cell.Value, 3) = "RAD" Or Left(cell.Value, 2) = "TB") And cell(1, 7).Value = 1.5 Then
        cell(1, 21).Value = ferulle_15
        End If
        If (Left(cell.Value, 3) = "RAD" Or Left(cell.Value, 2) = "TB") And cell(1, 7).Value = 2.5 Then
        cell(1, 21).Value = ferulle_25
        End If
        If (Left(cell.Value, 3) = "RAD" Or Left(cell.Value, 2) = "TB") And cell(1, 7).Value = 4 Then
        cell(1, 21).Value = ferulle_40
        End If

        
         If Left(cell.Value, 3) = "FCM" And cell(1, 13).Value = 1 And cell(1, 7).Value = 1 Then
        cell(1, 21).Value = ferulle_10
        End If
        If Left(cell.Value, 3) = "FCM" And cell(1, 13).Value = 1 And cell(1, 7).Value = 1.5 Then
        cell(1, 21).Value = ferulle_15
        End If
        If Left(cell.Value, 3) = "FCM" And cell(1, 13).Value = 1 And cell(1, 7).Value = 2.5 Then
        cell(1, 21).Value = ferulle_25
        End If
        If Left(cell.Value, 3) = "FCM" And cell(1, 13).Value = 1 And cell(1, 7).Value = 4 Then
        cell(1, 21).Value = ferulle_40
        End If
        
        

        
          If Left(cell.Value, 3) = "PFV" Then
        cell(1, 20).Value = 10
        End If
         If Left(cell.Value, 3) = "PFV" And cell(1, 7).Value = 1 Then
        cell(1, 20).Value = 11
        End If
        
        
          If (Left(cell.Value, 3) = "XDA" Or Left(cell.Value, 3) = "XDV") And cell(1, 7).Value = 1 Then
        cell(1, 21).Value = ferulle_10
        End If
        If (Left(cell.Value, 3) = "XDA" Or Left(cell.Value, 3) = "XDV") And cell(1, 7).Value = 1.5 Then
        cell(1, 21).Value = ferulle_15
        End If
        If (Left(cell.Value, 3) = "XDA" Or Left(cell.Value, 3) = "XDV") And cell(1, 7).Value = 2.5 Then
        cell(1, 21).Value = ferulle_25
        End If
        If (Left(cell.Value, 3) = "XDA" Or Left(cell.Value, 3) = "XDV") And cell(1, 7).Value = 4 Then
        cell(1, 21).Value = ferulle_40
        End If
        

                
        
            If (cell.Value = "K1" Or cell.Value = "K2" Or cell.Value = "K3" Or cell.Value = "K4") And cell(1, 7).Value = 1 Then
        cell(1, 21).Value = ferulle_10
        End If
        If (cell.Value = "K1" Or cell.Value = "K2" Or cell.Value = "K3" Or cell.Value = "K4") And cell(1, 7).Value = 1 Then
        cell(1, 21).Value = ferulle_15
        End If
        If (cell.Value = "K1" Or cell.Value = "K2" Or cell.Value = "K3" Or cell.Value = "K4") And cell(1, 7).Value = 1 Then
        cell(1, 21).Value = ferulle_25
        End If
        If (cell.Value = "K1" Or cell.Value = "K2" Or cell.Value = "K3" Or cell.Value = "K4") And cell(1, 7).Value = 1 Then
        cell(1, 21).Value = ferulle_40
        End If


        
         If (Left(cell.Value, 2) = "KA" Or Left(cell.Value, 3) = "RAA" Or Left(cell.Value, 3) = "XDS") And cell(1, 7).Value = 1 Then
        cell(1, 21).Value = ferulle_10
        End If
        If (Left(cell.Value, 2) = "KA" Or Left(cell.Value, 3) = "RAA" Or Left(cell.Value, 3) = "XDS") And cell(1, 7).Value = 1.5 Then
        cell(1, 21).Value = ferulle_15
        End If
        If (Left(cell.Value, 2) = "KA" Or Left(cell.Value, 3) = "RAA" Or Left(cell.Value, 3) = "XDS") And cell(1, 7).Value = 2.5 Then
        cell(1, 21).Value = ferulle_25
        End If
        If (Left(cell.Value, 2) = "KA" Or Left(cell.Value, 3) = "RAA" Or Left(cell.Value, 3) = "XDS") And cell(1, 7).Value = 4 Then
        cell(1, 21).Value = ferulle_40
        End If



         If Left(cell.Value, 2) = "KF" And cell(1, 7).Value = 1 Then
        cell(1, 21).Value = ferulle_10
        End If
        If Left(cell.Value, 2) = "KF" And cell(1, 7).Value = 1.5 Then
        cell(1, 21).Value = ferulle_15
        End If
        If Left(cell.Value, 2) = "KF" And cell(1, 7).Value = 2.5 Then
        cell(1, 21).Value = ferulle_25
        End If
        If Left(cell.Value, 2) = "KF" And cell(1, 7).Value = 4 Then
        cell(1, 21).Value = ferulle_40
        End If

                    
         If Left(cell.Value, 2) = "TF" And cell(1, 7).Value = 1 Then
        cell(1, 21).Value = ferulle_10
        End If
        If Left(cell.Value, 2) = "TF" And cell(1, 7).Value = 1.5 Then
        cell(1, 21).Value = ferulle_15
        End If
        If Left(cell.Value, 2) = "TF" And cell(1, 7).Value = 2.5 Then
        cell(1, 21).Value = ferulle_25
        End If
        If Left(cell.Value, 2) = "TF" And cell(1, 7).Value = 4 Then
        cell(1, 21).Value = ferulle_40
        End If
                    
                    
         If Left(cell.Value, 2) = "PF" And Not Left(cell.Value, 3) = "PFV" And cell(1, 7).Value = 1 Then
        cell(1, 21).Value = ferulle_10
        End If
        If Left(cell.Value, 2) = "PF" And Not Left(cell.Value, 3) = "PFV" And cell(1, 7).Value = 1.5 Then
        cell(1, 21).Value = ferulle_15
        End If
        If Left(cell.Value, 2) = "PF" And Not Left(cell.Value, 3) = "PFV" And cell(1, 7).Value = 2.5 Then
        cell(1, 21).Value = ferulle_25
        End If
        If Left(cell.Value, 2) = "PF" And Not Left(cell.Value, 3) = "PFV" And cell(1, 7).Value = 4 Then
        cell(1, 21).Value = ferulle_40
        End If



    '----------------------------Door Wireing ----------------------------
    
         If Left(cell.Value, 2) = "SF" And cell(1, 7).Value = 1 Then
        cell(1, 21).Value = ferulle_10
        End If
        If Left(cell.Value, 2) = "SF" And cell(1, 7).Value = 1.5 Then
        cell(1, 21).Value = ferulle_15
        End If
        If Left(cell.Value, 2) = "SF" And cell(1, 7).Value = 2.5 Then
        cell(1, 21).Value = ferulle_25
        End If
        If Left(cell.Value, 2) = "SF" And cell(1, 7).Value = 4 Then
        cell(1, 21).Value = ferulle_40
        End If
    
    
        
          If (Left(cell.Value, 3) = "SPM" Or Left(cell.Value, 3) = "STF") And cell(1, 7).Value = 1 Then
        cell(1, 21).Value = ferulle_10
        End If
        If (Left(cell.Value, 3) = "SPM" Or Left(cell.Value, 3) = "STF") And cell(1, 7).Value = 1.5 Then
        cell(1, 21).Value = ferulle_15
        End If
        If (Left(cell.Value, 3) = "SPM" Or Left(cell.Value, 3) = "STF") And cell(1, 7).Value = 2.5 Then
        cell(1, 21).Value = ferulle_25
        End If
        If (Left(cell.Value, 3) = "SPM" Or Left(cell.Value, 3) = "STF") And cell(1, 7).Value = 4 Then
        cell(1, 21).Value = ferulle_40
        End If
        

        If Left(cell.Value, 3) = "PGM" And cell(1, 7).Value = 1 Then
        cell(1, 21).Value = ferulle_10
        End If
        If Left(cell.Value, 3) = "PGM" And cell(1, 7).Value = 1.5 Then
        cell(1, 21).Value = ferulle_15
        End If
        If Left(cell.Value, 3) = "PGM" And cell(1, 7).Value = 2.5 Then
        cell(1, 21).Value = ferulle_25
        End If
        If Left(cell.Value, 3) = "PGM" And cell(1, 7).Value = 4 Then
        cell(1, 21).Value = ferulle_40
        End If
        
        
        If Left(cell.Value, 3) = "PG" And Not Left(cell.Value, 3) = "PGM" And cell(1, 7).Value = 1 Then
        cell(1, 21).Value = ferulle_10
        End If
        If Left(cell.Value, 3) = "PG" And Not Left(cell.Value, 3) = "PGM" And cell(1, 7).Value = 1.5 Then
        cell(1, 21).Value = ferulle_15
        End If
        If Left(cell.Value, 3) = "PG" And Not Left(cell.Value, 3) = "PGM" And cell(1, 7).Value = 2.5 Then
        cell(1, 21).Value = ferulle_25
        End If
        If Left(cell.Value, 3) = "PG" And Not Left(cell.Value, 3) = "PGM" And cell(1, 7).Value = 4 Then
        cell(1, 21).Value = ferulle_40
        End If
        

        
'---------------LOCKOUT RELAY---------------------
        
                If Left(cell.Value, 3) = "K86" And cell(1, 7).Value = 1 Then
        cell(1, 21).Value = ferulle_10
        End If
        If Left(cell.Value, 3) = "K86" And cell(1, 7).Value = 1.5 Then
        cell(1, 21).Value = ferulle_15
        End If
        If Left(cell.Value, 3) = "K86" And cell(1, 7).Value = 2.5 Then
        cell(1, 21).Value = ferulle_25
        End If
        If Left(cell.Value, 3) = "K86" And cell(1, 7).Value = 4 Then
        cell(1, 21).Value = ferulle_40
        End If
        


 '---------------------Wireing - 'Ref protection-----------------
 
          If cell.Value = "AA" And cell(1, 13).Value = 1 And ref.Range("J17").Value = "No" And cell(1, 7).Value = 1 Then
        cell(1, 21).Value = ferulle_10
        End If

          If cell.Value = "AA" And cell(1, 13).Value = 1 And ref.Range("J17").Value = "No" And cell(1, 7).Value = 1.5 Then
        cell(1, 21).Value = ferulle_15
        End If
       If cell.Value = "AA" And cell(1, 13).Value = 1 And ref.Range("J17").Value = "No" And cell(1, 7).Value = 2.5 Then
        cell(1, 21).Value = ferulle_25
        End If
        If cell.Value = "AA" And cell(1, 13).Value = 1 And ref.Range("J17").Value = "No" And cell(1, 7).Value = 4 Then
        cell(1, 21).Value = ferulle_40
        End If
        
        
                  If (Left(cell.Value, 2) = "BC" Or Left(cell.Value, 2) = "BE") And cell(1, 13).Value = 1 And cell(1, 7).Value = 1 Then
        cell(1, 21).Value = ferulle_10
        End If

                If (Left(cell.Value, 2) = "BC" Or Left(cell.Value, 2) = "BE") And cell(1, 13).Value = 1 And cell(1, 7).Value = 1.5 Then
        cell(1, 21).Value = ferulle_15
        End If
        If (Left(cell.Value, 2) = "BC" Or Left(cell.Value, 2) = "BE") And cell(1, 13).Value = 1 And cell(1, 7).Value = 2.5 Then
        cell(1, 21).Value = ferulle_25
        End If
        If (Left(cell.Value, 2) = "BC" Or Left(cell.Value, 2) = "BE") And cell(1, 13).Value = 1 And cell(1, 7).Value = 4 Then
        cell(1, 21).Value = ferulle_40
        End If
        
        
        
        If Left(cell.Value, 3) = "BAR" And cell(1, 7).Value = 1 Then
        cell(1, 21).Value = ferulle_10
        End If
        If Left(cell.Value, 3) = "BAR" And cell(1, 7).Value = 1.5 Then
        cell(1, 21).Value = ferulle_15
        End If
        If Left(cell.Value, 3) = "BAR" And cell(1, 7).Value = 2.5 Then
        cell(1, 21).Value = ferulle_25
        End If
        If Left(cell.Value, 3) = "BAR" And cell(1, 7).Value = 4 Then
        cell(1, 21).Value = ferulle_40
        End If






        
        
 End If
        
 Next
    
    
    
   Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True

End Sub


