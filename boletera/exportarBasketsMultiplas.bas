Attribute VB_Name = "exportarBasketsMultiplas"
Sub exportBasketMultiplas()
    ' -----
    ' Exporta baskets de "Ações" e "Ações - Múltiplos"

    ' Acionada pelo botão "EXPORTAR BASKET (XP PO)"
    ' na planilha "DASH BSKT MÚLTIPLAS" para as pastas
    ' "Baskets" e "Receita" no mesmo local da boletera
    ' em "ONEDRIVE/Ferramentas/Boletera"
    ' -----

    'Melhora performance desligando oeracoes do excel
    Application.ScreenUpdating = False

    'Importa variáveis globais
    EstaPastaDeTrabalho.Importar_Variaveis_Globais

    Dim arqBoletera As Workbook
    Dim wbRece As Workbook

    Dim boletera As Worksheet
    Dim basket As Worksheet
    Dim export As Worksheet
    Dim dash As Worksheet
    Dim base As Worksheet

    Dim valor As String
    Dim endereco As String
    Dim data As String
    Dim cliente As String
    Dim broker As String
    Dim nome As String
    Dim caminho_pasta_baskets As String
    Dim caminho_pasta_receita_broker As String
    Dim caminho_arquivo_modelo_receita_avulsa As String
    Dim caminho_arquivo_receita_avulsa_broker As String

    Dim i As Integer
    Dim Dlin As Integer
    Dim ultlin As Integer

    Dim testeNA As Boolean

    Set arqBoletera = ThisWorkbook
    Set boletera = arqBoletera.Sheets("BOLET. ORDENS MÚLTIPLAS")
    Set basket = arqBoletera.Sheets("BASKET - MÚLTIPLAS")
    Set export = arqBoletera.Sheets("EXPORT BSKT MÚLTIPLAS")
    Set dash = arqBoletera.Sheets("DASH BSKT MÚLTIPLAS")
    Set base = arqBoletera.Sheets("BASE")

    caminho_pasta_boletera = EstaPastaDeTrabalho.fso.BuildPath(EstaPastaDeTrabalho.ONEDRIVE_GERAL, "Ferramentas\Boletera")
    caminho_pasta_baskets = EstaPastaDeTrabalho.fso.BuildPath(caminho_pasta_boletera, "Baskets")
    caminho_pasta_receita_broker = EstaPastaDeTrabalho.fso.BuildPath(caminho_pasta_boletera, "Receita\" & boletera.Range("H5").Value)
    caminho_arquivo_receita_avulsa_broker = EstaPastaDeTrabalho.fso.BuildPath(caminho_pasta_receita_broker, "RECEITA AVULSA.xlsx")
    caminho_arquivo_modelo_receita_avulsa = EstaPastaDeTrabalho.fso.BuildPath(caminho_pasta_boletera, "Templates\RECEITA AVULSA.xlsx")

    'Cria a pasta de receita do broker se não existir
    'com o arquivo modelo de receita avulsa
    If Dir(caminho_pasta_receita_broker, vbDirectory) = "" Then
        MkDir (caminho_pasta_receita_broker)
        FileCopy caminho_arquivo_modelo_receita_avulsa, caminho_arquivo_receita_avulsa_broker
    End If

    data = CStr(base.Range("AK8").Value)
    cliente = base.Range("AL7")
    nome = dash.Range("C5").Value

    On Error Resume Next

    broker = base.Range("as6").Value
    codbroker = base.Range("AM6").Value

    cont = 1
    For i = 6 To 5000
        ultlin = dash.Cells(i, 3).Row
        valor = dash.Cells(i, 3).Value
        If valor = "" Then
            Exit For
        ElseIf (cont < 5) Then
            nome = nome & "_" & valor
            cont = cont + 1
        End If
    Next
    ultlin = ultlin - 1

    'Se cliente EX Mesa, registra no nome do arquivo como "NOVO"
    On Error Resume Next
    testeNA = (boletera.Range("C5").Value = CVErr(xlErrNA))
    If testeNA Then
        nome = "NOVO"
    End If

    nome_arquivo_acoes_multiplos = "(AÇÕES - MÚLTIPLOS) " & Year(Date) & " " & Format(Month(Date), "00") & " " & Format(Day(Date), "00") & " " & nome & " " & boletera.Range("h5").Value & " " & codbroker & ".xlsx"

    export.Range("A1:R500").AutoFilter
    export.Range("A1:R500").AutoFilter
    export.Range("A1:R1", export.Range("A1:R1").End(xlDown)).Copy

    Workbooks.Add
    ActiveSheet.Paste
    ActiveWorkbook.SaveAs Filename := EstaPastaDeTrabalho.fso.BuildPath(caminho_pasta_baskets, nome_arquivo_acoes_multiplos), FileFormat:=xlOpenXMLWorkbook, CreateBackup := False
    ActiveWindow.Close

    'Separando as baskets
    For i = 5 To 5000
        valor = dash.Cells(i, 3).Value

        If valor = "" Then
            Exit For
        End If

        nome = valor
        On Error Resume Next
        testeNA = (boletera.Range("C5").Value = CVErr(xlErrNA))
        If testeNA Then
            nome = "NOVO"
        End If

        On Error Resume Next
        If dash.Range("D" & i).Value <> "" Then
            nome_arquivo_acoes = "(AÇÕES) " & Year(Date) & " " & Format(Month(Date), "00") & " " & Format(Day(Date), "00") & " " & nome & " " & boletera.Range("H5").Value & " " & base.Range("AM6").Value & ".xlsx"

            export.Range("A1:R500").AutoFilter Field:=1, Criteria1:=nome
            export.Range("A1:R1", export.Range("A1:R1").End(xlDown)).Copy

            Workbooks.Add
            ActiveSheet.Paste
            ActiveWorkbook.SaveAs Filename := EstaPastaDeTrabalho.fso.BuildPath(caminho_pasta_baskets, nome_arquivo_acoes), FileFormat := xlOpenXMLWorkbook, CreateBackup := False
            ActiveWindow.Close
        End If

        wbRece = Workbooks.Open(caminho_arquivo_receita_avulsa_broker, Password := "2022")

        Dlin = Range("A1").End(xlDown).Row + 1

        If Dlin > 200000 Then
            Dlin = 2
        End If

        Range("A" & Dlin).Value = data
        Range("B" & Dlin).Value = valor
        Range("C" & Dlin).Value = broker

        ActiveWorkbook.Close savechanges:=True
    Next

    'Reativa configuraçõs do excel
    Application.ScreenUpdating = True
End Sub
