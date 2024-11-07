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
Attribute EXPORT_BASKET.VB_ProcData.VB_Invoke_Func = " \n14"

    'macro ok
    
'
' EXPORT_BASKET Macro
'
    Dim arqBoletera As Workbook
    Dim boletera As Worksheet
    Dim basket As Worksheet
    Dim export As Worksheet
    Dim dash As Worksheet
    Dim base As Worksheet
'
    Dim endereco As String
    Dim salvar As String
    Dim data As String
    Dim cliente As String
    Dim broker
    Dim testestr
    Dim strpath
    Dim endesave
    Dim salvarBoleta
    Dim nome As String
    Dim testeNA
    Dim arqnome
    Dim wbRece
    Dim Dlin
    
    
    Set arqBoletera = ThisWorkbook
    Set boletera = arqBoletera.Sheets("BOLET. AVULSAS")
    Set basket = arqBoletera.Sheets("BASKET")
    Set export = arqBoletera.Sheets("EXPORT BSKT")
    Set dash = arqBoletera.Sheets("DASH BSKT")
    Set base = arqBoletera.Sheets("BASE")
    Set bullpro = arqBoletera.Sheets("BULL PRO")
    

    'Application.ScreenUpdating = False

    ChDir ThisWorkbook.Path
    
    strpath = ThisWorkbook.Path
    For i = 65 To 90
        letra = Chr(i)
        If Left(strpath, 1) = letra Then
            ChDrive letra
            GoTo proximaparte
        End If
    Next
proximaparte:
    ' Sobe dois níveis para a raiz do operacional
    ChDir ".."
    endesave2 = CurDir
    ChDir ".."
    endesave = CurDir

        
    
    salvar = endesave2 & "\3 - RECEITA\" & boletera.Range("F5").Value
    salvarBoleta = endesave & "\0 - AÇÕES\2 - BASKETS\"
    modelo = endesave2 & "\3 - RECEITA\MODELO\RECEITA AVULSA.xlsx"
    If (Dir(salvar, vbDirectory) = "") Then
        MkDir (salvar)
        FileCopy modelo, salvar & "\RECEITA AVULSA.xlsx"
    End If
    salvar = salvar & "\RECEITA AVULSA.xlsx"
    data = CStr(base.Range("AK8").Value)
    cliente = base.Range("AL7")
    On Error Resume Next
    broker = base.Range("AM7").Value
    nome = boletera.Range("C5").Value
    
    
    
    'Se cliente EX Mesa, registra no nome do arquivo como "NOVO"
    On Error Resume Next
    testeNA = (boletera.Range("C5").Value = CVErr(xlErrNA))
    If (testeNA) Then
        nome = "NOVO"
    End If
    
    arqnome = "(AÇÕES) " & Year(Date) & " " & Format(Month(Date), "00") & " " & Format(Day(Date), "00") & " " & nome & " " & boletera.Range("C4").Value & " " & boletera.Range("F5") & " " & broker
    
    'Sheets("EXPORT BSKT").Select
    'Range("A1:R1").Select
    'Range(Selection, Selection.End(xlDown)).Select
    'Selection.Copy
    'Range("A1").Select
    
    export.Range("A1:R1", export.Range("A1:R1").End(xlDown)).Copy
    
    Workbooks.Add
    ActiveSheet.Paste
    MsgBox salvarBoleta & arqnome & ".xlsx"
    ActiveWorkbook.SaveAs Filename:= _
        (salvarBoleta & arqnome & ".xlsx") _
        , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWindow.Close
    
    bullpro.Range("A1:R1", bullpro.Range("A1:R1").End(xlDown)).Copy
    
    Workbooks.Add
    ActiveSheet.Paste
    ActiveSheet.Range("A1:R1", ActiveSheet.Range("A1:R1").End(xlDown)).Copy
    ActiveSheet.Range("A1").PasteSpecial xlPasteValues
    'ActiveSheet.PasteSpecial xlPasteValues
    MsgBox salvarBoleta & arqnome & ".csv"
    ActiveWorkbook.SaveAs Filename:= _
        (salvarBoleta & arqnome & ".csv") _
        , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWindow.Close
    
    If dash.Range("D5").Value <> "" Then GoTo skip
    On Error Resume Next
    wbRece = Workbooks.Open(salvar, Password:="2022")
    
    Dlin = Range("A1").End(xlDown).Row + 1
    
    If Dlin > 200000 Then
        Dlin = 2
    End If
    
    Range("A" & Dlin).Value = data
    Range("B" & Dlin).Value = cliente
    Range("C" & Dlin).Value = broker
    ActiveWorkbook.Save
    ActiveWorkbook.Close
skip:
    Application.ScreenUpdating = True
    
    'MsgBox "Basket exportada"
    
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
Attribute LIMPAR_BASKET.VB_ProcData.VB_Invoke_Func = " \n14"

    
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
