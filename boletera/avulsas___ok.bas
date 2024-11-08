Attribute VB_Name = "avulsas___ok"
Sub GRAVAR_BASKET()

    'Macro ok

    Dim arqBoletera As Workbook
    Dim boletera As Worksheet
    Dim basket As Worksheet
    Dim export As Worksheet
    Dim dash As Worksheet

    Dim ultlin As Integer
    Dim i As Integer
    Dim valor
    Dim answer As Integer

    Set arqBoletera = ThisWorkbook
    Set boletera = arqBoletera.Sheets("BOLET. AVULSAS")
    Set basket = arqBoletera.Sheets("BASKET")
    Set export = arqBoletera.Sheets("EXPORT BSKT")
    Set dash = arqBoletera.Sheets("DASH BSKT")
    Set base = arqBoletera.Sheets("BASE")

    fim = base.Range("AU7").End(xlDown).Row

    cliente = WorksheetFunction.CountIf(base.Range("AU7:AU" & fim), boletera.Range("C4").Value)

    ultlinha = boletera.Range("B10").End(xlDown).Row


basket:
    ultlin = 2
'    Application.ScreenUpdating = False

    basket.Range("a3:R150").Copy

    For i = 2 To 5000
        ultlin = export.Cells(i, 1).Row
        valor = export.Cells(i, 1).Value
        If (valor = "") Then
            Exit For
        End If
    Next

    export.Range("A" & ultlin).PasteSpecial xlPasteValues
    On Error Resume Next
    With export.AutoFilter.Sort
        .SortFields.Clear
        .SortFields.Add2 Key _
            :=Range("A1:A10000"), SortOn:=xlSortOnValues, Order:=xlAscending, _
            DataOption:=xlSortNormal
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    On Error GoTo 0
    For i = 2 To 5000
        ultlin = export.Cells(i, 1).Row
        valor = export.Cells(i, 1).Value
        If (valor = "") Then
            Exit For
        End If
    Next
    ultlin = ultlin - 1

    export.Range("A2:A" & ultlin).Copy


    If IsEmpty(dash.Range("c5")) Then
        dash.Range("c5").PasteSpecial xlPasteValues, skipblanks:=True

    Else
        dash.Range("c4").End(xlDown).Offset(1).PasteSpecial xlPasteValues
    End If

    dash.Range("$C4:$C1000").RemoveDuplicates Columns:=1, Header:=xlYes

    Application.ScreenUpdating = True
    dash.Select

    'MsgBox "Basket gravada"

End Sub
Sub EXPORT_BASKET()

    ' Importa variáveis globais, incluindo ONEDRIVE_GERAL
    EstaPastaDeTrabalho.Importar_Variaveis_Globais

    ' Declaração de variáveis
    Dim arqBoletera As Workbook
    Dim boletera As Worksheet
    Dim basket As Worksheet
    Dim export As Worksheet
    Dim dash As Worksheet
    Dim base As Worksheet
    Dim bullpro As Worksheet
    Dim fso As Object
    Dim caminhoReceita As String
    Dim caminhoBaskets As String
    Dim caminhoModelo As String
    Dim salvar As String
    Dim salvarBoleta As String
    Dim modelo As String
    Dim data As String
    Dim cliente As String
    Dim broker As Variant
    Dim nome As String
    Dim testeNA As Boolean
    Dim arqnome As String
    Dim wbRece As Workbook
    Dim Dlin As Long

    ' Define os objetos da pasta de trabalho e as planilhas
    Set arqBoletera = ThisWorkbook
    Set boletera = arqBoletera.Sheets("BOLET. AVULSAS")
    Set basket = arqBoletera.Sheets("BASKET")
    Set export = arqBoletera.Sheets("EXPORT BSKT")
    Set dash = arqBoletera.Sheets("DASH BSKT")
    Set base = arqBoletera.Sheets("BASE")
    Set bullpro = arqBoletera.Sheets("BULL PRO")

    ' Cria o objeto FileSystemObject para manipulação de caminhos
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Define os caminhos principais usando a variável global ONEDRIVE_GERAL
    caminhoReceita = fso.BuildPath(ONEDRIVE_GERAL, "Ferramentas\Boletera\Receita")
    caminhoBaskets = fso.BuildPath(ONEDRIVE_GERAL, "Ferramentas\Boletera\Baskets")
    caminhoModelo = fso.BuildPath(caminhoReceita, "MODELO/RECEITA AVULSA.xlsx")

    ' Define os caminhos de salvamento e nome do arquivo de modelo
    salvar = fso.BuildPath(caminhoReceita, boletera.Range("F5").Value)
    salvarBoleta = caminhoBaskets
    modelo = caminhoModelo

    ' Cria diretório e copia o arquivo modelo caso não exista
    If (Dir(salvar, vbDirectory) = "") Then
        MkDir (salvar)
        FileCopy modelo, fso.BuildPath(salvar, "RECEITA AVULSA.xlsx")
    End If
    salvar = fso.BuildPath(salvar, "RECEITA AVULSA.xlsx")

    ' Define valores de data, cliente, broker e nome
    data = CStr(base.Range("AK8").Value)
    cliente = base.Range("AL7")
    broker = base.Range("AM7").Value
    nome = boletera.Range("C5").Value

    ' Se cliente for "EX Mesa", registra como "NOVO" no nome do arquivo
    testeNA = (boletera.Range("C5").Value = CVErr(xlErrNA))
    If (testeNA) Then
        nome = "NOVO"
    End If

    ' Define o nome do arquivo a ser salvo
    arqnome = "(AÇÕES) " & Year(Date) & " " & Format(Month(Date), "00") & " " & Format(Day(Date), "00") & " " & nome & " " & boletera.Range("C4").Value & " " & boletera.Range("F5") & " " & broker

    ' Exporta a planilha "EXPORT BSKT" para um novo arquivo Excel
    export.Range("A1:R1", export.Range("A1:R1").End(xlDown)).Copy
    Workbooks.Add
    ActiveSheet.Paste
    ActiveWorkbook.SaveAs Filename:=fso.BuildPath(salvarBoleta, arqnome & ".xlsx"), FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWindow.Close

    ' Exporta a planilha "BULL PRO" para um arquivo CSV
    bullpro.Range("A1:R1", bullpro.Range("A1:R1").End(xlDown)).Copy
    Workbooks.Add
    ActiveSheet.Paste
    ActiveSheet.Range("A1:R1", ActiveSheet.Range("A1:R1").End(xlDown)).Copy
    ActiveSheet.Range("A1").PasteSpecial xlPasteValues
    ActiveWorkbook.SaveAs Filename:=fso.BuildPath(salvarBoleta, arqnome & ".csv"), FileFormat:=xlCSV, CreateBackup:=False
    ActiveWindow.Close

    ' Verifica se há dados na célula D5 da planilha "DASH BSKT"
    If dash.Range("D5").Value <> "" Then GoTo skip

    ' Abre o arquivo "RECEITA AVULSA.xlsx" e insere os dados
    On Error Resume Next
    Set wbRece = Workbooks.Open(salvar, Password:="2022")

    Dlin = Range("A1").End(xlDown).Row + 1
    If Dlin > 200000 Then Dlin = 2

    Range("A" & Dlin).Value = data
    Range("B" & Dlin).Value = cliente
    Range("C" & Dlin).Value = broker
    ActiveWorkbook.Save
    ActiveWorkbook.Close

skip:
    Application.ScreenUpdating = True

End Sub

Sub EXPORT_BASKET_BULL()

    Dim bull As Worksheet
    ThisWorkbook.Unprotect Password:="senhadaboletera"
    On Error Resume Next
    ThisWorkbook.Sheets("BULL").Visible = True
    ThisWorkbook.Sheets("BULL PRO").Visible = True
' macro ok
    Set bull = ThisWorkbook.Sheets("BULL")
    Call EXPORT_BASKET

    i = 1
    Do While Not bull.Cells(i, 1).Value = "":
        i = i + 1
    Loop
    bull.Range("A2:E" & i - 1).Copy

    ThisWorkbook.Protect Structure:=True, Windows:=False, Password:="senhadaboletera"
    ThisWorkbook.Sheets("BULL").Select
    'MsgBox ("Basket copiada")
    ' Workbooks(arq).Close SaveChanges:=True
End Sub

Sub EXPORT_BASKET_XPCIO()

    ThisWorkbook.Unprotect Password:="senhadaboletera"
    On Error Resume Next
    ThisWorkbook.Sheets("EXPORT XP").Visible = True
    Dim arqBoletera As Workbook
    Dim basket As Worksheet
    Dim export As Worksheet
    Dim exportXP As Worksheet
    Dim i As Integer
    Dim lRow

    Call EXPORT_BASKET

    Set arqBoletera = ThisWorkbook
    Set basket = arqBoletera.Sheets("BASKET")
    Set export = arqBoletera.Sheets("EXPORT BSKT")
    Set exportXP = arqBoletera.Sheets("EXPORT XP")

    export.Range("A2:A81").Copy
    exportXP.Range("C3").PasteSpecial Paste:=xlPasteValues
    export.Range("C2", "C81").Copy
    exportXP.Range("D3").PasteSpecial Paste:=xlPasteValues
    export.Range("D2", "D81").Copy
    exportXP.Range("E3").PasteSpecial Paste:=xlPasteValues
    export.Range("B2", "B81").Copy
    exportXP.Range("G3").PasteSpecial Paste:=xlPasteValues
    export.Range("E2", "E81").Copy
    exportXP.Range("F3").PasteSpecial Paste:=xlPasteValues

    i = 1
    Do While Not IsEmpty(exportXP.Cells(i, 1).Value):
        i = i + 1
    Loop

    lRow = exportXP.Cells(Rows.Count, 3).End(xlUp).Row
    ThisWorkbook.Sheets("EXPORT XP").Select
    exportXP.Range("A1:E" & i - 1).Copy
    ThisWorkbook.Protect Structure:=True, Windows:=False, Password:="senhadaboletera"
    'MsgBox "Basket copiada"
End Sub

Sub LIMPAR_BASKET()


    Sheets("EXPORT BSKT").Range("A2:R1048576").ClearContents
    Range("C5:C1048576").ClearContents
    Sheets("EXPORT XP").Range("C3:G100").ClearContents
    Sheets("TWAP CIO").Range("C3:K100").ClearContents

End Sub

Sub AtualizaCotacao()

'macro ok

Dim lRow As Long

lRow = Cells(Rows.Count, 3).End(xlUp).Row
Range("C7:C" & lRow).Value = Range("C7:C" & lRow).FormulaR1C1

'MsgBox "Feito"

End Sub

Sub BASKET_TWAP_CIO()

    ThisWorkbook.Unprotect Password:="senhadaboletera"
    On Error Resume Next
    ThisWorkbook.Sheets("TWAP CIO").Visible = True
    Dim arqBoletera As Workbook
    Dim basket As Worksheet
    Dim export As Worksheet
    Dim exportXP As Worksheet
    Dim i As Integer
    Dim lRow

    Call EXPORT_BASKET

    Set arqBoletera = ThisWorkbook
    Set basket = arqBoletera.Sheets("BASKET")
    Set export = arqBoletera.Sheets("EXPORT BSKT")
    Set exportXP = arqBoletera.Sheets("TWAP CIO")

    export.Range("A2:A81").Copy
    exportXP.Range("C3").PasteSpecial Paste:=xlPasteValues

    export.Range("C2", "C81").Copy
    exportXP.Range("D3").PasteSpecial Paste:=xlPasteValues

    export.Range("D2", "D81").Copy
    exportXP.Range("E3").PasteSpecial Paste:=xlPasteValues

    export.Range("B2", "B81").Copy
    exportXP.Range("F3").PasteSpecial Paste:=xlPasteValues

    export.Range("E2", "E81").Copy
    exportXP.Range("H3").PasteSpecial Paste:=xlPasteValues

    i = 1
    Do While Not IsEmpty(exportXP.Cells(i, 1).Value):
        i = i + 1
    Loop

    lRow = exportXP.Cells(Rows.Count, 3).End(xlUp).Row
    ThisWorkbook.Sheets("TWAP CIO").Select
    exportXP.Range("A1:K" & i - 1).Copy
    ThisWorkbook.Protect Structure:=True, Windows:=False, Password:="senhadaboletera"
    'MsgBox "Basket copiada"
End Sub
