Attribute VB_Name = "Módulo7"
Function FormatarEmail(caminhoArquivo As String, dados As Object) As String
    ' Funcao para abrir e formatar arquivo de template de email com codificação UTF-8
    ' Args:
    '     - caminhoArquivo: caminho do arquivo como path para ser aberto e formatado
    '     - dados: dicionário contendo as informações para formatar o arquivo
    ' Raises:
    '     - 13: Type Mismatch, o argumento não é um dicionário
    '     - 75: Path/File access error, não foi possível ler o arquivo
    ' Returns:
    '     - Texto do arquivo aberto e formatado

    ' Declara argumento para key do dict
    Dim chave As Variant
    Dim conteudoArquivo As String
    Dim stream As Object
    
    ' Checa se o argumento dados é Dicionário
    If Not TypeName(dados) = "Dictionary" Then
        Err.Raise Number:=13, Description:="Type Mismatch: O argumento não é um dicionário válido."
    End If
    
    ' Criar o objeto ADODB.Stream para leitura em UTF-8
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' Tipo de dados de texto
    stream.Charset = "UTF-8" ' Configuração para UTF-8
    stream.Open
    stream.LoadFromFile caminhoArquivo

    ' Ler o conteúdo do arquivo e armazenar na variável
    conteudoArquivo = stream.ReadText
    stream.Close
    
    'Removendo variavel da memoria
    Set stream = Nothing

    ' Formatar o conteúdo do arquivo
    For Each chave In dados.Keys
        ' Substituir cada placeholder pelo valor correspondente
        conteudoArquivo = Replace(conteudoArquivo, "{{" & chave & "}}", dados(chave))
    Next chave

    ' Return da função
    FormatarEmail = conteudoArquivo
End Function

Function IdentificarCaminhoTemplate(valor As String) As String
    ' Função para indentificar o template referente ao tipo de estrutura
    ' Args:
        '- Valor: tipo de estrutura escolhido pelo broker
        
    ' Returns:
        '- O caminho do arquivo HTML

    Dim userProfilePath As String
    Set fs = CreateObject("Scripting.FileSystemObject")

    userProfilePath = Environ("USERPROFILE")
    CaminhoTemplates = fs.BuildPath(userProfilePath, "XP Investimentos\Manchester - Mesa RV - Backoffice - Backoffice\Projetos\Envio de email estruturadas boletera\templates")

    Select Case valor
        Case "Alocação Protegida"
            IdentificarCaminhoTemplate = fs.BuildPath(CaminhoTemplates, "alocacaoprotegida.html")
        
        Case "Booster"
            IdentificarCaminhoTemplate = fs.BuildPath(CaminhoTemplates, "booster.html")
            
        Case "Booster Shield"
            IdentificarCaminhoTemplate = fs.BuildPath(CaminhoTemplates, "boostershield.html")
            
        Case "Collar UI"
            IdentificarCaminhoTemplate = fs.BuildPath(CaminhoTemplates, "collarui.html")
            
        Case "Financiamento"
            IdentificarCaminhoTemplate = fs.BuildPath(CaminhoTemplates, "financiamento.html")
            
        Case "Rubi"
            IdentificarCaminhoTemplate = fs.BuildPath(CaminhoTemplates, "rubi.html")
        
        Case Else
            IdentificarCaminhoTemplate = ""
    End Select
End Function
Function FormatarParaMoeda(valor As Double) As String
    ' Funcao que formata valores para tipo Moeda
    ' Args:
        '- valor: valor que sera formatado
    
    ' Returns:
        '- valor formatado

    FormatarParaMoeda = "R$ " & Format(valor, "#,##0.00")
    
End Function

Function CriarDict(nomeEstrutura As String, ws As Worksheet) As Object
    'Funcao para criar dicionario com dados de uma estrutura com
    'base no nome da estrutura
    
    'Args:
        '- nomeEstrutura: nome da na célula D11 usada para criar o dicionário
        '- ws: nome do sheets do arquivo Excel
    'Returns:
        '- dicionário com os argumentos dos arquivos HTML de cada estrutura

    'Declarando o dicionario
    
    Dim dados As Object
    Set dados = CreateObject("Scripting.Dictionary")

    ' Arrumar a conta do financeiro exposto quant x preço

    Select Case nomeEstrutura
        Case "Alocação Protegida"
            dados.Add "Conta", ws.Range("A11").Value
            dados.Add "Ativo", ws.Range("J11").Value
            dados.Add "Compra", ws.Range("H11").Value
            dados.Add "Quantidade", ws.Range("K11").Value
            dados.Add "strike", Format(ws.Range("L11").Value, "0.00%")
            dados.Add "premio", Format(ws.Range("M11").Value, "0.00%")
            dados.Add "Preco", FormatarParaMoeda(ws.Range("N11").Value)
            dados.Add "Vencimento", Format(ws.Range("O11").Value, "dd/mm/yyyy")
            dados.Add "operacao", ws.Range("P11").Value
            dados.Add "FinanceiroExposto", FormatarParaMoeda(ws.Range("K11").Value * ws.Range("N11").Value)
            dados.Add "PerdaMax", Format(ws.Range("L11").Value - ws.Range("M11").Value, "0.00%")
        
        Case "Booster"
            dados.Add "Conta", ws.Range("A11").Value
            dados.Add "Ativo", ws.Range("J11").Value
            dados.Add "Compra", ws.Range("H11").Value
            dados.Add "Quantidade", ws.Range("K11").Value
            dados.Add "PrecoRef", FormatarParaMoeda(ws.Range("L11").Value)
            dados.Add "Vencimento", Format(ws.Range("M11").Value, "dd/mm/yyyy")
            dados.Add "StrikeCallVendida", Format(ws.Range("N11").Value, "0.00%")
            dados.Add "StrikeCallComprada", Format(ws.Range("O11").Value, "0.00%")
            dados.Add "operacao", ws.Range("P11").Value
            dados.Add "FinanceiroExposto", FormatarParaMoeda(ws.Range("K11").Value * ws.Range("L11").Value)
            dados.Add "SubtracaoStrike", Format((ws.Range("N11").Value - 1), "0.00%")
            dados.Add "GanhoMax", Format((ws.Range("N11").Value - 1 + ws.Range("N11").Value - ws.Range("O11").Value), "0.00%")

        Case "Booster Shield"
            dados.Add "Conta", ws.Range("A11").Value
            dados.Add "Ativo", ws.Range("J11").Value
            dados.Add "Compra", ws.Range("H11").Value
            dados.Add "Quantidade", ws.Range("K11").Value
            dados.Add "PrecoRef", FormatarParaMoeda(ws.Range("L11").Value)
            dados.Add "Vencimento", Format(ws.Range("M11").Value, "dd/mm/yyyy")
            dados.Add "StrikePutComprada", Format(ws.Range("N11").Value, "0.00%")
            dados.Add "StrikeCallVendida", Format(ws.Range("O11").Value, "0.00%")
            dados.Add "StrikeCallComprada", Format(ws.Range("P11").Value, "0.00%")
            dados.Add "Barreira", Format(ws.Range("Q11").Value, "0.00%")
            dados.Add "operacao", Format(ws.Range("R11").Value, "0.00%")
            dados.Add "BarreiraMilUm", Format((ws.Range("Q11").Value - 1.0001), "0.00%")
            dados.Add "SubtracaoBarreira", Format((ws.Range("Q11").Value - 1), "0.00%")
            dados.Add "SubtracaoStrike", Format((ws.Range("O11").Value - 1), "0.00%")
            dados.Add "FinanceiroExposto", FormatarParaMoeda(ws.Range("K11").Value * ws.Range("L11").Value)
            dados.Add "GanhoMax", Format(2 * ws.Range("Q11").Value - ws.Range("P11").Value - 1, "0.00%")
            dados.Add "PerdaMax", Format(ws.Range("N11").Value - 1, "0.00%")

        Case "Collar UI"
            dados.Add "Conta", ws.Range("A11").Value
            dados.Add "Ativo", ws.Range("J11").Value
            dados.Add "Compra", ws.Range("H11").Value
            dados.Add "Quantidade", ws.Range("K11").Value
            dados.Add "Preco", FormatarParaMoeda(ws.Range("L11").Value)
            dados.Add "Vencimento", Format(ws.Range("M11").Value, "dd/mm/yyyy")
            dados.Add "StrikePut", Format(ws.Range("N11").Value, "0.00%")
            dados.Add "StrikeCall", Format(ws.Range("O11").Value, "0.00%")
            dados.Add "Barreira", Format(ws.Range("P11").Value, "0.00%")
            dados.Add "operacao", ws.Range("Q11").Value
            dados.Add "FinanceiroExposto", FormatarParaMoeda(ws.Range("K11").Value * ws.Range("L11").Value)
            dados.Add "SubtracaoBarreira", Format((ws.Range("P11").Value - 1), "0.00%")
            dados.Add "BarreiraMilUm", Format((ws.Range("P11").Value - 1.0001), "0.00%")
            dados.Add "SubtracaoStrikeCall", Format((ws.Range("O11").Value - 1), "0.00%")
            dados.Add "SubtracaoStrikePut", Format((1 - ws.Range("N11").Value), "0.00%")

        Case "Financiamento"
            dados.Add "Conta", ws.Range("A11").Value
            dados.Add "Ativo", ws.Range("J11").Value
            dados.Add "Compra", ws.Range("H11").Value
            dados.Add "Quantidade", ws.Range("K11").Value
            dados.Add "Preco", FormatarParaMoeda(ws.Range("L11").Value)
            dados.Add "Vencimento", Format(ws.Range("M11").Value, "dd/mm/yyyy")
            dados.Add "strike", Format(ws.Range("N11").Value, "0.00%")
            dados.Add "premio", Format(ws.Range("O11").Value, "0.00%")
            dados.Add "operacao", ws.Range("P11").Value
            dados.Add "SubtracaoStrike", Format((ws.Range("N11").Value - 1), "0.00%")
            dados.Add "FinanceiroExposto", FormatarParaMoeda(ws.Range("K11").Value * ws.Range("L11").Value)
            dados.Add "GanhoMax", Format(ws.Range("N11").Value + ws.Range("O11").Value, "0.00%")

        Case "Rubi"
            dados.Add "Conta", ws.Range("A11").Value
            dados.Add "Ativo", ws.Range("J11").Value
            dados.Add "Compra", ws.Range("H11").Value
            dados.Add "Quantidade", ws.Range("K11").Value
            dados.Add "PrecoRef", FormatarParaMoeda(ws.Range("L11").Value)
            dados.Add "Vencimento", Format(ws.Range("M11").Value, "dd/mm/yyyy")
            dados.Add "Strike", Format(ws.Range("N11").Value, "0.00%")
            dados.Add "Barreira", Format(ws.Range("O11").Value, "0.00%")
            dados.Add "operacao", ws.Range("P11").Value
            dados.Add "SubtracaoBarreira", Format((1 - ws.Range("O11").Value), "0.00%")
            dados.Add "FinanceiroExposto", FormatarParaMoeda(ws.Range("K11").Value * ws.Range("L11").Value)
            dados.Add "B1", Format((SubtracaoBarreira + 0.0001), "0.00%")
            dados.Add "SubtracaoStrike", Format((ws.Range("N11").Value - 1), "0.00%")
        
    End Select

    Set CriarDict = dados
End Function

Function montarEmail(emailAssessor As String, emailCliente As String, corpoEmail As String, assuntoEmail As String)
    ' Função que monta o email de envio das operações
    
    'Args:
        '- emailAssessor: Email que ficará em cópia
        '- emailCliente: Email de quem receberá o email
        '- corpoEmail: Email

    Dim outApp As Outlook.Application
    Dim OutMail As Outlook.MailItem
    
    Set outApp = CreateObject("Outlook.Application")
    Set OutMail = outApp.CreateItem(olMailItem)

    With OutMail
        .Subject = assuntoEmail
        .HTMLBody = corpoEmail
        .CC = emailAssessor
        .To = emailCliente
        .Display
    End With

End Function

Sub EnviarEmailOperação()
    Dim caminhoTemplate As String
    Dim dadosEmail As Object
    Dim corpoEmail As String
    Dim assuntoEmail As String
    Dim emailAssessor As String
    Dim emailCliente As String
    Dim nomeEstrutura As String
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Sheets("ENVIO OP. ESTRUTRADAS")
    
    ' Definindo valores para as variáveis
    nomeEstrutura = ws.Range("G11").Value
    emailAssessor = ws.Range("E11").Value
    emailCliente = ws.Range("C11").Value
    assuntoEmail = "Detalhes da Estrutura - " & nomeEstrutura
    
    ' Identificar o caminho do template
    caminhoTemplate = IdentificarCaminhoTemplate(nomeEstrutura)
    
    ' Criar o dicionário com os dados do email
    Set dadosEmail = CriarDict(nomeEstrutura, ws)
    
    ' Formatar o corpo do email com o template
    corpoEmail = FormatarEmail(caminhoTemplate, dadosEmail)
    
    ' Montar e enviar o email
    montarEmail emailAssessor, emailCliente, corpoEmail, assuntoEmail
End Sub
