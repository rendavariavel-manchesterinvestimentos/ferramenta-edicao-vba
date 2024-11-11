Attribute VB_Name = "M�dulo6"
Sub limpar_solicita��o_offshore()

Set aba_offshore = ThisWorkbook.Sheets("OFFSHORE")
aba_offshore.Range("A11:A50").Cells.ClearContents
aba_offshore.Range("D11:D50").Cells.ClearContents
aba_offshore.Range("E11:E50").Cells.ClearContents
End Sub

Sub exportar_solicita��o_offshore()

Set aba_offshore = ThisWorkbook.Sheets("OFFSHORE")

broker = aba_offshore.Range("B7")
If IsEmpty(Range("B7").Value) Then
        MsgBox "Broker n�o selecionado, selecione um broker na c�lula B7.", vbExclamation, "Aviso"
        Exit Sub
    End If

Set planilha_nova = Workbooks.Add
Set aba_nova = planilha_nova.Sheets(1)

aba_offshore.Range("A10:G50").Copy
aba_nova.Range("A1").PasteSpecial xlPasteValues

planilha_nova.SaveAs Filename:=ThisWorkbook.Path & "\Baskets offshore\" & broker & " - " & Format(Now, "yyyy-mm-dd_hh-mm-ss") & ".xlsx"
planilha_nova.Close savechanges:=True
MsgBox "Solicita��o enviada com sucesso!", , "Tudo certo"
End Sub
