Attribute VB_Name = "EmailAuditoria"
Sub emailOrdemTermo()
Dim outApp As Outlook.Application
Dim OutMail As Outlook.MailItem
Dim OutAccount As Outlook.Account
Dim strbody As String
Dim EnviarPara As String
Dim Mensagem As String


For i = 13 To 1048576
    If Sheets("ORDENS MESA").Range("AF" & i).Value = "" Then GoTo fim
Next
fim:


    Set outApp = CreateObject("Outlook.Application")
    Set OutMail = outApp.CreateItem(olMailItem)
    Mensagem = RangetoHTML(Sheets("ORDENS MESA").Range("AF8:AL" & i - 1))
    OutMail.Display
    On Error Resume Next
    With OutMail
        .Subject = "Auditoria para execução de ordens - Manchester/XP"
        .HTMLBody = Mensagem & "<br>" & .HTMLBody
        .Display
    End With
    On Error GoTo 0
    Set OutMail = Nothing
    Set outApp = Nothing
    Call EXPORT_BASKET

End Sub
Sub emailOrdemMercado()
Dim outApp As Outlook.Application
Dim OutMail As Outlook.MailItem
Dim OutAccount As Outlook.Account
Dim strbody As String
Dim EnviarPara As String
Dim Mensagem As String


For i = 13 To 1048576
    If Sheets("ORDENS MESA").Range("Z" & i).Value = "" Then GoTo fim
Next
fim:


    Set outApp = CreateObject("Outlook.Application")
    Set OutMail = outApp.CreateItem(olMailItem)
    Mensagem = RangetoHTML(Sheets("ORDENS MESA").Range("Z8:AD" & i - 1))
    OutMail.Display
    On Error Resume Next
    With OutMail
        .Subject = "Auditoria para execução de ordens - Manchester/XP"
        .HTMLBody = Mensagem & "<br>" & .HTMLBody
        .Display
    End With
    On Error GoTo 0
    Set OutMail = Nothing
    Set outApp = Nothing
    Call EXPORT_BASKET

End Sub
Sub emailOrdemPreco()
Dim outApp As Outlook.Application
Dim OutMail As Outlook.MailItem
Dim OutAccount As Outlook.Account
Dim strbody As String
Dim EnviarPara As String
Dim Mensagem As String


For i = 13 To 1048576
    If Sheets("ORDENS MESA").Range("S" & i).Value = "" Then GoTo fim
Next
fim:


    Set outApp = CreateObject("Outlook.Application")
    Set OutMail = outApp.CreateItem(olMailItem)
    OutMail.Display
    Mensagem = RangetoHTML(Sheets("ORDENS MESA").Range("S8:X" & i - 1))
    Signature = OutMail.HTMLBody
    On Error Resume Next
    With OutMail
        .Subject = "Auditoria para execução de ordens - Manchester/XP"
        .HTMLBody = Mensagem & Signature
        .Display
    End With
    On Error GoTo 0
    Set OutMail = Nothing
    Set outApp = Nothing
    Call EXPORT_BASKET

End Sub
Sub emailOrdemCIO_merc()
Dim outApp As Outlook.Application
Dim OutMail As Outlook.MailItem
Dim OutAccount As Outlook.Account
Dim strbody As String
Dim EnviarPara As String
Dim Mensagem As String


For i = 13 To 1048576
    If Sheets("ORDENS MESA").Range("AN" & i).Value = "" Then GoTo fim
Next
fim:


    Set outApp = CreateObject("Outlook.Application")
    Set OutMail = outApp.CreateItem(olMailItem)
    Mensagem = RangetoHTML(Sheets("ORDENS MESA").Range("AN8:AS" & i - 1))
    OutMail.Display
    On Error Resume Next
    With OutMail
        .Subject = "Auditoria para execução de ordens - Manchester/XP"
        .HTMLBody = Mensagem & "<br>" & .HTMLBody
        .Display
    End With
    On Error GoTo 0
    Set OutMail = Nothing
    Set outApp = Nothing
    Call EXPORT_BASKET

End Sub
Sub emailOrdemCIO_preco()
Dim outApp As Outlook.Application
Dim OutMail As Outlook.MailItem
Dim OutAccount As Outlook.Account
Dim strbody As String
Dim EnviarPara As String
Dim Mensagem As String


For i = 13 To 1048576
    If Sheets("ORDENS MESA").Range("AU" & i).Value = "" Then GoTo fim
Next
fim:


    Set outApp = CreateObject("Outlook.Application")
    Set OutMail = outApp.CreateItem(olMailItem)
    Mensagem = RangetoHTML(Sheets("ORDENS MESA").Range("AU8:AZ" & i - 1))
    OutMail.Display
    On Error Resume Next
    With OutMail
        .Subject = "Auditoria para execução de ordens - Manchester/XP"
        .HTMLBody = Mensagem & "<br>" & .HTMLBody
        .Display
    End With
    On Error GoTo 0
    Set OutMail = Nothing
    Set outApp = Nothing
    Call EXPORT_BASKET

End Sub
Function RangetoHTML(rng As Range)

    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook

    TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

    'Copy the range and create a new workbook to past the data in
    rng.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=1
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
    RangetoHTML = ts.readall
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

