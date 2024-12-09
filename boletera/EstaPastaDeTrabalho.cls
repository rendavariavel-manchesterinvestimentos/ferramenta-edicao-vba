'Initialize the variables
Public fso As Object
Public USERPROFILE As String
Public ONEDRIVE_GERAL As String

Private Sub Workbook_Open()
    'Funcao para ser executada na abertura do arquivo, usamos
    'para inicializar as variaveis globais do projeto

    'Refs:
        'https://support.microsoft.com/pt-br/office/executar-automaticamente-uma-macro-ao-abrir-uma-pasta-de-trabalho-1e55959b-e077-4c88-a696-c3017600db44

    'Objeto para trabalhar com caminhos
    Set fso = CreateObject("Scripting.FileSystemObject")

    USERPROFILE = Environ("USERPROFILE")

    'Arquivos na cloud
    ONEDRIVE_GERAL = fso.BuildPath(USERPROFILE, "XP Investimentos\Manchester - Mesa RV - General")

End Sub

