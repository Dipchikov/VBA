Attribute VB_Name = "Mail"
Sub Mail_Selection_Range_Outlook_Body()
'For Tips see: http://www.rondebruin.nl/win/winmail/Outlook/tips.htm
'Don't forget to copy the function RangetoHTML in the module.
'Working in Excel 2000-2016
    Dim rng As Range
    Dim OutApp As Object
    Dim OutMail As Object
    Dim strbody As String
    Dim Signature As String

    
    Set rng = Nothing
    On Error Resume Next
    'Only the visible cells in the selection
    'Set rng = Selection.SpecialCells(xlCellTypeVisible)
    Set rng = Application.Selection
        lr = Range("A" & Rows.Count).End(xlUp).Row

    Set rng = Range("A6:E" & lr + 3).SpecialCells(xlCellTypeVisible)

 'GET DEFAULT EMAIL SIGNATURE
    Signature = Environ("appdata") & "\Microsoft\Signatures\"
    If Dir(Signature, vbDirectory) <> vbNullString Then
        Signature = Signature & Dir$(Signature & "*.htm")
    Else:
        Signature = ""
    End If
    Signature = CreateObject("Scripting.FileSystemObject").GetFile(Signature).OpenAsTextStream(1, -2).ReadAll
    
'Set rng = Application.InputBox("Range1:", xTitleId, rng.Address, Type:=8)
    'You can also use a fixed range if you want
    'Set rng = Sheets("YourSheet").Range("D4:D12").SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    If rng Is Nothing Then
        MsgBox "The selection is not a range or the sheet is protected" & _
               vbNewLine & "please correct and try again.", vbOKOnly
        Exit Sub
         
    End If

    With Application
        .EnableEvents = False
        .ScreenUpdating = False
    End With


    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    On Error Resume Next
    With OutMail
        .To = Range("B1").Value
        .CC = Range("B2").Value
        .BCC = ""
        .Subject = Range("B3").Value
        .HTMLBody = RangetoHTML(rng) & Signature
        .Display   'or use .Send
    End With
    On Error GoTo 0

    With Application
        .EnableEvents = True
        .ScreenUpdating = True
    End With

    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub


Function RangetoHTML(rng As Range)
' Changed by Ron de Bruin 28-Oct-2006
' Working in Office 2000-2016
    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook
    Dim lr As Long
    TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

    'Copy the range and create a new workbook to past the data in
    rng.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With

    'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         Filename:=TempFile, _
         Sheet:=TempWB.Sheets(1).Name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With

    'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.ReadAll
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")



    'Close TempWB
    TempWB.Close savechanges:=False

    'Delete the htm file we used in this function
    Kill TempFile

    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
End Function

