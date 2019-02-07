Sub Komax_table_inter()
    '-----------Table columns----------------------
    Dim JobKey As String
    Dim JobNumber As String
    Dim JobTotalPiecess As Integer
    Dim JobBatchSizes As Integer
    Dim JobName As String
    Dim JobHint As String
    Dim ArticleKey As String
    Dim ArticleGroup As String
    Dim ArticleName As String
    Dim ArticleBundlingSide As Single
    Dim LeadSetStrippingLength1 As String
    Dim LeadSetStrippingLength2 As String
    Dim LeadSetStrippingLength3 As String
    Dim LeadSetTerminalKey1 As String
    Dim LeadSetTerminalKey2 As String
    Dim LeadSetTerminalKey3 As String
    Dim MarkingTextBegin1_2_Text1 As String
    Dim MarkingTextBegin1_2_Text2 As String
    Dim MarkingTextBegin1_2_Text3 As String
    Dim MarkingTextBegin1_2_Turns As Integer
    Dim MarkingTextEnd1_2_Text1 As String
    Dim MarkingTextEnd1_2_Text2 As String
    Dim MarkingTextEnd1_2_Text3 As String
    Dim MarkingTextEnd1_2_Turn As Integer
    Dim MarkingTextEndless1_2_Text As String
    Dim MarkingTextEndless1_2_Turn As Integer
    Dim BundlingPostProcess As Integer

    Dim ferulle_10 As String
    Dim ferulle_15 As String
    Dim ferulle_25 As String
    Dim ferulle_40 As String




    Dim lrdata As Long
    Dim lrFinal As Long

Set Data = Sheets("Interconnections")
Set Final = Sheets("Komax")
Set Base = Sheets("Type of cables")

    If Komax_inter.ferulles_Yes.Value = True Then

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


    '------Komax table columns-----------------------------------------------------
    JobTotalPiecess = 1
    JobBatchSizes = 1
    JobKey = Left(Data.Range("D1").Value, 10) & "W" & Right(Data.Range("D1").Value, 4)
    ArticleKey = JobKey
    JobName = "WA for " & Data.Range("D1").Value
    ArticleName = JobName
    ArticleGroup = "Italy\UniSec\" & Right(Data.Range("B1").Value, 4) & "####"
    MarkingTextBegin1_2_Turns = 0
    MarkingTextEnd1_2_Turn = 1
    MarkingTextEndless1_2_Turn = 1

    If Komax_inter.Bundling_one.Value = True Then
        ArticleBundlingSide = 1
    Else
        ArticleBundlingSide = 3
    End If


    lrdata = Range("A" & Rows.Count).End(xlUp).Row


    '----- Clar Komax table-----------------

    Final.Range("A2:CO1048576").EntireRow.Delete
    '----------Prigram number------------------

    ' Number_pr_comax.number

    '------------------------------------------
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    On Error Resume Next

    '-------------------------------------------

    Data.Active
    Data.ShowAllData

    Final.Range("A2").Select


    For i = 15 To lrdata
        '-------ïúðâè ðåä â final  range--------------------
        j = i - 13

        MarkingTextBegin1_2_Text1 = Data.Range("C" & i).Value
        MarkingTextBegin1_2_Text2 = MarkingTextBegin1_2_Text1
        MarkingTextBegin1_2_Text3 = MarkingTextBegin1_2_Text1
        MarkingTextEnd1_2_Text1 = Data.Range("F" & i).Value
        MarkingTextEnd1_2_Text2 = MarkingTextEnd1_2_Text1
        MarkingTextEnd1_2_Text3 = MarkingTextEnd1_2_Text1
        MarkingTextEndless1_2_Text = MarkingTextEnd1_2_Text1
        If Not Data.Range("G" & i).Value = 4 Then
            LeadSetStrippingLength1 = 10
            LeadSetStrippingLength2 = 10
        Else
            LeadSetStrippingLength1 = 12
            LeadSetStrippingLength2 = 12
        End If
        '----------ferules-- -----------------


        If Komax_inter.ferulles_Yes.Value = True Then

            If Data.Range("G" & i).Value = 4 And ferulle_40 <> "" Then
                LeadSetTerminalKey1 = ferulle_40
                LeadSetTerminalKey2 = ferulle_40
                Final.Range("AA" & j).Value = LeadSetTerminalKey1
                Final.Range("AB" & j).Value = LeadSetTerminalKey2
            Else
                If Data.Range("G" & i).Value = 4 Then
                    LeadSetStrippingLength1 = 12
                    LeadSetStrippingLength2 = 12
                    Final.Range("O" & j).Value = LeadSetStrippingLength1
                    Final.Range("P" & j).Value = LeadSetStrippingLength2
                End If
            End If

            If Data.Range("G" & i).Value = 2.5 And ferulle_25 <> "" Then
                LeadSetTerminalKey1 = ferulle_25
                LeadSetTerminalKey2 = ferulle_25
                Final.Range("AA" & j).Value = LeadSetTerminalKey1
                Final.Range("AB" & j).Value = LeadSetTerminalKey2
            Else
                If Data.Range("G" & i).Value = 2.5 Then
                    LeadSetStrippingLength1 = 10
                    LeadSetStrippingLength2 = 10
                    Final.Range("O" & j).Value = LeadSetStrippingLength1
                    Final.Range("P" & j).Value = LeadSetStrippingLength2
                End If
            End If

            If Data.Range("G" & i).Value = 1.5 And ferulle_15 <> "" Then
            LeadSetTerminalKey1 = ferulle_15
            LeadSetTerminalKey2 = ferulle_15
            Final.Range("AA" & j).Value = LeadSetTerminalKey1
            Final.Range("AB" & j).Value = LeadSetTerminalKey2
        Else
            If Data.Range("G" & i).Value = 1.5 Then
                LeadSetStrippingLength1 = 10
                LeadSetStrippingLength2 = 10
                Final.Range("O" & j).Value = LeadSetStrippingLength1
                Final.Range("P" & j).Value = LeadSetStrippingLength2
            End If
        End If

        If Data.Range("G" & i).Value = 1 And ferulle_10 <> "" Then
            LeadSetTerminalKey1 = ferulle_10
            LeadSetTerminalKey2 = ferulle_10
            Final.Range("AA" & j).Value = LeadSetTerminalKey1
            Final.Range("AB" & j).Value = LeadSetTerminalKey2
        Else
            If Data.Range("G" & i).Value = 1 Then
                LeadSetStrippingLength1 = 10
                LeadSetStrippingLength2 = 10
                Final.Range("O" & j).Value = LeadSetStrippingLength1
                Final.Range("P" & j).Value = LeadSetStrippingLength2
            End If
        End If

        End If

        If Komax_inter.ferulles_Yes.Value = False Then

            If Data.Range("G" & i).Value = 1 Then
                LeadSetStrippingLength1 = 10
                LeadSetStrippingLength2 = 10
                Final.Range("O" & j).Value = LeadSetStrippingLength1
                Final.Range("P" & j).Value = LeadSetStrippingLength2
            End If

            If Data.Range("G" & i).Value = 1.5 Then
                LeadSetStrippingLength1 = 10
                LeadSetStrippingLength2 = 10
                Final.Range("O" & j).Value = LeadSetStrippingLength1
                Final.Range("P" & j).Value = LeadSetStrippingLength2
            End If

            If Data.Range("G" & i).Value = 2.5 Then
                LeadSetStrippingLength1 = 10
                LeadSetStrippingLength2 = 10
                Final.Range("O" & j).Value = LeadSetStrippingLength1
                Final.Range("P" & j).Value = LeadSetStrippingLength2
            End If
            If Data.Range("G" & i).Value = 4 Then
                LeadSetStrippingLength1 = 12
                LeadSetStrippingLength2 = 12
                Final.Range("O" & j).Value = LeadSetStrippingLength1
                Final.Range("P" & j).Value = LeadSetStrippingLength2
            End If

        End If


        '----------Condition If cell is empty------------------

        If Not (Data.Range("L" & i).Value = "-" Or Data.Range("L" & i).Value = "Shielded cable") Then
            '-----------------------Äåôèíèðàíå íà ïðîãðàìà ïîä 99 ðåäà-----------------
            Final.Range("C" & j).Value = JobTotalPiecess
            Final.Range("D" & j).Value = JobBatchSizes
            Final.Range("I" & j).Value = Final.Range("E" & j).Value
            Final.Range("M" & j).Value = Data.Range("I" & i).Value
            Final.Range("K" & j).Value = Data.Range("J" & i).Value

            ' Final.Range("O" & j).Value = LeadSetStrippingLength1
            'Final.Range("P" & j).Value = LeadSetStrippingLength2

            Final.Range("AG" & j).Value = MarkingTextBegin1_2_Text1
            Final.Range("AH" & j).Value = MarkingTextBegin1_2_Text2
            Final.Range("AI" & j).Value = MarkingTextBegin1_2_Text2
            Final.Range("AJ" & j).Value = MarkingTextBegin1_2_Turns
            Final.Range("AK" & j).Value = MarkingTextEnd1_2_Text1
            Final.Range("AL" & j).Value = MarkingTextEnd1_2_Text2
            Final.Range("AM" & j).Value = MarkingTextEnd1_2_Text3
            Final.Range("AO" & j).Value = MarkingTextEndless1_2_Text
            Final.Range("AN" & j).Value = MarkingTextEnd1_2_Turn
            Final.Range("AP" & j).Value = MarkingTextEndless1_2_Turn
            Final.Range("BA" & j).Value = ArticleBundlingSide

            '-------------------------cut bandling-----------------------------

            If Komax_inter.Bundling_Cut_Yes = True Then
                If Left(Data.Range("A" & i).Value, 4) <> Left(Data.Range("A" & i + 1).Value, 4) And i <> lrdata Then
                    Final.Range("BC" & j).Value = 2
                Else
                    Final.Range("BC" & j).Value = 1
                End If
            End If

            If Komax_inter.Bundling_Cut_No = True Then
                Final.Range("BC" & j).Value = 1
            End If





        End If
    Next i

    '----------Condition If cell is empty------------------
    Final.Range("C:C").SpecialCells(xlCellTypeBlanks).EntireRow.Delete



    Dim number As Integer
    'Dim k As Long
    'Dim j As Long
    Dim sum As Long


    sum = 1
    lrFinal = Final.Range("C" & Rows.Count).End(xlUp).Row

    number = (lrFinal) / 99 + 1

    For k = 1 To number
        For l = 1 To 99
            sum = sum + 1
            If lrFinal <= 100 Then
                Final.Range("A" & sum).Value = JobKey
                Final.Range("G" & sum).Value = ArticleKey
            Else
                Final.Range("A" & sum).Value = JobKey & "." & k
                Final.Range("G" & sum).Value = ArticleKey & "." & k
            End If
            Final.Range("E" & sum).Value = JobName
            Final.Range("I" & sum).Value = ArticleName
            Final.Range("H" & sum).Value = ArticleGroup
            If sum = lrFinal Then
                Exit For
            End If
        Next l
        If sum = lrFinal Then
            Exit For
        End If
    Next k



    Sheets("nterconnections").Select
    Range("A15").Select



    Dim wb As Workbook
    Set wb = Workbooks.Add
    ThisWorkbook.Sheets("Komax").Copy Before:=wb.Sheets(1)



        '-------------add user in Footer ---------------
    With ActiveSheet.PageSetup
        .LeftFooter = "&D" & Chr(13) & "&9" & Application.UserName
        .RightFooter = "Page " & "&P" & Chr(13) & "&9" & Tools.Label8.Caption
    End With

    '---------Èçòðèâàíå íà Sheet1------------------

    Application.DisplayAlerts = False
    Sheets("Sheet1").Delete
    Application.DisplayAlerts = True


    Application.CutCopyMode = False 'esp

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True


    Dim sFileSaveName As Variant
    Dim sPath As String
    sPath = "Inter_" & Right(Data.Range("B1").Value, 4) & "_" & Data.Range("D1").Value
    InitialFoldr$ = "\\10.28.38.5\ppmv\Productions\Italian\LVC\UniSec\!!!__Orders\!_____Ongoing Orders"
    sFileSaveName = Application.GetSaveAsFilename(InitialFileName:=Left(sPath, 26), FileFilter:="Excel Files (*.csv), *.xlsm")
    If sFileSaveName <> False Then
        ActiveWorkbook.SaveAs sFileSaveName, FileFormat:=xlCSV, Local:=True
Application.DisplayAlerts = False
        ActiveWorkbook.Close
        Application.DisplayAlerts = True
    End If



End Sub


