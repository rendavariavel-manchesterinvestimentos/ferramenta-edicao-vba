Attribute VB_Name = "M�dulo8"
Sub atualizar_estrutura()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("TESTE")

    If ws.Range("G11").Value = "Aloca��o Protegida" Then
        ws.Range("J10").Value = "ATIVO"
        ws.Range("K10").Value = "QUANTIDADE"
        ws.Range("L10").Value = "STRIKE"
        ws.Range("M10").Value = "PR�MIO"
        ws.Range("N10").Value = "PRE�O"
        ws.Range("O10").Value = "VENCIMENTO"
        ws.Range("P10").Value = "OPERA��O"
        ws.Range("Q10").Value = " "
        ws.Range("R10").Value = " "

    ElseIf ws.Range("G11").Value = "Booster" Then
        ws.Range("J10").Value = "ATIVO"
        ws.Range("K10").Value = "QUANTIDADE"
        ws.Range("L10").Value = "PRE�O REF"
        ws.Range("M10").Value = "VENCIMENTO"
        ws.Range("N10").Value = "STRIKE CALL VENDIDA"
        ws.Range("O10").Value = "STRIKE CALL COMPRADA"
        ws.Range("P10").Value = "OPERA��O"
        ws.Range("Q10").Value = " "
        ws.Range("R10").Value = " "

    ElseIf ws.Range("G11").Value = "Booster Shield" Then
        ws.Range("J10").Value = "ATIVO"
        ws.Range("K10").Value = "QUANTIDADE"
        ws.Range("L10").Value = "PRE�O REF"
        ws.Range("M10").Value = "VENCIMENTO"
        ws.Range("N10").Value = "STRIKE PUT COMPRADA"
        ws.Range("O10").Value = "STRIKE CALL VENDIDA"
        ws.Range("P10").Value = "STRIKE CALL COMPRADA"
        ws.Range("Q10").Value = "BARREIRA"
        ws.Range("R10").Value = "OPERA��O"

    ElseIf ws.Range("G11").Value = "Collar UI" Then
        ws.Range("J10").Value = "ATIVO"
        ws.Range("K10").Value = "QUANTIDADE"
        ws.Range("L10").Value = "PRE�O"
        ws.Range("M10").Value = "VENCIMENTO"
        ws.Range("N10").Value = "STRIKE PUT"
        ws.Range("O10").Value = "STRIKE CALL"
        ws.Range("P10").Value = "BARREIRA"
        ws.Range("Q10").Value = "OPERA��O"
        ws.Range("R10").Value = " "

    ElseIf ws.Range("G11").Value = "Financiamento" Then
        ws.Range("J10").Value = "ATIVO"
        ws.Range("K10").Value = "QUANTIDADE"
        ws.Range("L10").Value = "PRE�O"
        ws.Range("M10").Value = "VENCIMENTO"
        ws.Range("N10").Value = "STRIKE"
        ws.Range("O10").Value = "PR�MIO"
        ws.Range("P10").Value = "OPERA��O"
        ws.Range("Q10").Value = " "
        ws.Range("R10").Value = " "

    ElseIf ws.Range("G11").Value = "NDF" Then
        ws.Range("J10").Value = "PRE�O COMPRA"
        ws.Range("K10").Value = "PRE�O REF"
        ws.Range("L10").Value = "VENCIMENTO"
        ws.Range("M10").Value = "VOLUME"
        ws.Range("N10").Value = "DATA"
        ws.Range("O10").Value = "OPERA��O"
        ws.Range("P10").Value = " "
        ws.Range("Q10").Value = " "
        ws.Range("R10").Value = " "

    ElseIf ws.Range("G11").Value = "NDF com CAP" Then
        ws.Range("J10").Value = "PRE�O COMPRA"
        ws.Range("K10").Value = "PRE�O REF"
        ws.Range("L10").Value = "VENCIMENTO"
        ws.Range("M10").Value = "VOLUME"
        ws.Range("N10").Value = "DATA"
        ws.Range("O10").Value = "OPERA��O"
        ws.Range("P10").Value = "CAP"
        ws.Range("Q10").Value = " "
        ws.Range("R10").Value = " "

    ElseIf ws.Range("G11").Value = "Rubi" Then
        ws.Range("J10").Value = "ATIVO"
        ws.Range("K10").Value = "QUANTIDADE"
        ws.Range("L10").Value = "PRE�O REF"
        ws.Range("M10").Value = "VENCIMENTO"
        ws.Range("N10").Value = "STRIKE"
        ws.Range("O10").Value = "BARREIRA"
        ws.Range("P10").Value = "OPERA��O"
        ws.Range("Q10").Value = " "
        ws.Range("R10").Value = " "

    Else
        MsgBox "A estrutura n�o foi definida"

    End If
End Sub

Sub LimparConteudo()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("TESTE")

    ' Limpa o conte�do do intervalo de G11 at� R11
    ws.Range("G11:R11").ClearContents

    ws.Range("A11").ClearContents

End Sub
