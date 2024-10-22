Attribute VB_Name = "exemple"
Function FormatTextWithDict(caminhoArquivo As String, dados As Object) As String
    'Funcao para abrir um arquivo de texto "template"
    'Args:
        '- caminhoArquivo: caminho do arquivo como path para ser aberto e formatado
        '- dados: dicionario contendo as informacoes para formatar o arquivo
    'Raises:
        '- 13: Type Mismatch, o argumento não é um dicionário
        '- 75: Path/File access error, não foi possível lêr o arquivo
    'Returns:
        '- Texto do arquivo aberto e formatado

    Dim chave As Variant
    Dim numeroArquivo As Integer
    Dim conteudoArquivo As String

    'Checa se o argumento Dados é Dicionário
    If Not TypeName(dados) = "Dictionary" Then
        Err.Raise Number:=13, Description:="Type Mismatch: O argumento não é um dicionário válido."
    End If

    'Informa qual espaco de memória usar
    numeroArquivo = FreeFile

    'Abrir o arquivo para leitura
    On Error GoTo ErroAbrirArquivo
    Open caminhoArquivo For Input As numeroArquivo
        'LOF lê o arquivo em bytes e aloca na variável
        conteudoArquivo = Input(LOF(numeroArquivo), numeroArquivo)
    Close numeroArquivo

    'Formatar o arquivo
    For Each chave In dados.Keys
        'Substituir cada placeholder pelo valor correspondente
        conteudoArquivo = Replace(conteudoArquivo, "{{" & chave & "}}", dados(chave))
    Next chave

    'Return da função
    FormatTextWithDict = conteudoArquivo
    Exit Function

'Tratamento de erro ao ler arquivo
ErroAbrirArquivo:
    Err.Raise Number:=75, Description:="Path/File access error: Não foi possível abrir o arquivo"
End Function
