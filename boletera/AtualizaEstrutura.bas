Attribute VB_Name = "AtualizaEstrutura"
Sub atualizar_estrutura()
    'Atualiza os valores das colunas de acordo com a estrutura escolhida pelo Broker
    'Utiliza o botão Atualizar Estrutura para rodar o codigo

    Dim ws As Worksheet
    'Faz a formatação de cada uma das estruturas
    'Muda o valor da coluna e o título de cada uma delas
    Set ws = ThisWorkbook.Sheets("ENVIO OP. ESTRUTRADAS")

    If ws.Range("G11").Value = "Alocação Protegida" Then
        ws.Columns("J:J").NumberFormat = "@"
        ws.Range("J11").HorizontalAlignment = xlRight
        ws.Range("J10").Value = "ATIVO"
        Columns("K:K").NumberFormat = "0"
        ws.Range("K10").Value = "QUANTIDADE"
        ws.Columns("L:L").NumberFormat = "0.00%"
        ws.Range("L10").Value = "STRIKE"
        ws.Columns("M:M").NumberFormat = "0.00%"
        ws.Range("M10").Value = "PRÊMIO"
        ws.Columns("N:N").NumberFormat = "$ #,##0.00"
        ws.Range("N10").Value = "PREÇO DA AÇÃO"
        ws.Columns("O:O").NumberFormat = "dd/mm/yyyy"
        ws.Range("O10").Value = "VENCIMENTO"
        ws.Range("P11").HorizontalAlignment = xlRight
        ws.Range("P10").Value = "TIPO DE OPERAÇÃO"
        ws.Range("P11").Value = ws.Range("G11")
        ws.Range("Q10").Value = " "
        ws.Range("R10").Value = " "
        ws.Range("Q11").Value = " "
        ws.Range("R11").Value = " "

    ElseIf ws.Range("G11").Value = "Booster" Then
        ws.Columns("J:J").NumberFormat = "@"
        ws.Range("J11").HorizontalAlignment = xlRight
        ws.Range("J10").Value = "ATIVO"
        Columns("K:K").NumberFormat = "0"
        ws.Range("K10").Value = "QUANTIDADE"
        ws.Columns("L:L").NumberFormat = "$ #,##0.00"
        ws.Range("L10").Value = "PREÇO DA AÇÃO"
        ws.Columns("M:M").NumberFormat = "dd/mm/yyyy"
        ws.Range("M10").Value = "VENCIMENTO"
        ws.Columns("N:N").NumberFormat = "0.00%"
        ws.Range("N10").Value = "STRIKE CALL VENDIDA"
        ws.Columns("O:O").NumberFormat = "0.00%"
        ws.Range("O10").Value = "STRIKE CALL COMPRADA"
        ws.Range("P11").HorizontalAlignment = xlRight
        ws.Range("P10").Value = "TIPO DE OPERAÇÃO"
        ws.Range("P11").Value = ws.Range("G11")
        ws.Range("Q10").Value = " "
        ws.Range("R10").Value = " "
        ws.Range("Q11").Value = " "
        ws.Range("R11").Value = " "

    ElseIf ws.Range("G11").Value = "Booster Shield" Then
        ws.Columns("J:J").NumberFormat = "@"
        ws.Range("J11").HorizontalAlignment = xlRight
        ws.Range("J10").Value = "ATIVO"
        Columns("K:K").NumberFormat = "0"
        ws.Range("K10").Value = "QUANTIDADE"
        ws.Columns("L:L").NumberFormat = "$ #,##0.00"
        ws.Range("L10").Value = "PREÇO DA AÇÃO"
        ws.Columns("M:M").NumberFormat = "dd/mm/yyyy"
        ws.Range("M10").Value = "VENCIMENTO"
        ws.Columns("N:N").NumberFormat = "0.00%"
        ws.Range("N10").Value = "STRIKE PUT COMPRADA"
        ws.Columns("O:O").NumberFormat = "0.00%"
        ws.Range("O10").Value = "STRIKE CALL VENDIDA"
        ws.Columns("P:P").NumberFormat = "0.00%"
        ws.Range("P10").Value = "STRIKE CALL COMPRADA"
        ws.Columns("Q:Q").NumberFormat = "0.00%"
        ws.Range("Q10").Value = "BARREIRA"
        ws.Range("R11").HorizontalAlignment = xlRight
        ws.Range("R10").Value = "TIPO DE OPERAÇÃO"
        ws.Range("R11").Value = ws.Range("G11")

    ElseIf ws.Range("G11").Value = "Collar UI" Then
        ws.Columns("J:J").NumberFormat = "@"
        ws.Range("J11").HorizontalAlignment = xlRight
        ws.Range("J10").Value = "ATIVO"
        Columns("K:K").NumberFormat = "0"
        ws.Range("K10").Value = "QUANTIDADE"
        ws.Columns("L:L").NumberFormat = "$ #,##0.00"
        ws.Range("L10").Value = "PREÇO DA AÇÃO"
        ws.Columns("M:M").NumberFormat = "dd/mm/yyyy"
        ws.Range("M10").Value = "VENCIMENTO"
        ws.Columns("N:N").NumberFormat = "0.00%"
        ws.Range("N10").Value = "STRIKE PUT"
        ws.Columns("O:O").NumberFormat = "0.00%"
        ws.Range("O10").Value = "STRIKE CALL"
        ws.Columns("P:P").NumberFormat = "0.00%"
        ws.Range("P10").Value = "BARREIRA"
        ws.Range("Q11").HorizontalAlignment = xlRight
        ws.Range("Q10").Value = "TIPO DE OPERAÇÃO"
        ws.Range("Q11").Value = ws.Range("G11")
        ws.Range("R10").Value = " "
        ws.Range("R11").Value = " "

    ElseIf ws.Range("G11").Value = "Financiamento" Then
        ws.Columns("J:J").NumberFormat = "@"
        ws.Range("J11").HorizontalAlignment = xlRight
        ws.Range("J10").Value = "ATIVO"
        Columns("K:K").NumberFormat = "0"
        ws.Range("K10").Value = "QUANTIDADE"
        ws.Columns("L:L").NumberFormat = "$ #,##0.00"
        ws.Range("L10").Value = "PREÇO DA AÇÃO"
        ws.Columns("M:M").NumberFormat = "dd/mm/yyyy"
        ws.Range("M10").Value = "VENCIMENTO"
        ws.Columns("N:N").NumberFormat = "0.00%"
        ws.Range("N10").Value = "STRIKE"
        ws.Columns("O:O").NumberFormat = "0.00%"
        ws.Range("O10").Value = "PRÊMIO"
        ws.Range("P11").HorizontalAlignment = xlRight
        ws.Range("P10").Value = "TIPO DE OPERAÇÃO"
        ws.Range("P11").Value = ws.Range("G11")
        ws.Range("Q10").Value = " "
        ws.Range("R10").Value = " "
        ws.Range("Q11").Value = " "
        ws.Range("R11").Value = " "

    ElseIf ws.Range("G11").Value = "Rubi" Then
        ws.Columns("J:J").NumberFormat = "@"
        ws.Range("J11").HorizontalAlignment = xlRight
        ws.Range("J10").Value = "ATIVO"
        Columns("K:K").NumberFormat = "0"
        ws.Range("K10").Value = "QUANTIDADE"
        ws.Columns("L:L").NumberFormat = "$ #,##0.00"
        ws.Range("L10").Value = "PREÇO DA AÇÃO"
        ws.Columns("M:M").NumberFormat = "dd/mm/yyyy"
        ws.Range("M10").Value = "VENCIMENTO"
        ws.Columns("N:N").NumberFormat = "0.00%"
        ws.Range("N10").Value = "STRIKE"
        ws.Columns("O:O").NumberFormat = "0.00%"
        ws.Range("O10").Value = "BARREIRA"
        ws.Range("P11").HorizontalAlignment = xlRight
        ws.Range("P10").Value = "TIPO DE OPERAÇÃO"
        ws.Range("P11").Value = ws.Range("G11")
        ws.Range("Q10").Value = " "
        ws.Range("R10").Value = " "
        ws.Range("Q11").Value = " "
        ws.Range("R11").Value = " "

    Else
        MsgBox "A estrutura não foi definida"

    End If
End Sub

Sub LimparConteudo()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("ENVIO OP. ESTRUTRADAS")

    ' Limpa o conteúdo do intervalo de G11 até R11
    ws.Range("G11:R11").ClearContents

    ws.Range("A11").ClearContents

End Sub
