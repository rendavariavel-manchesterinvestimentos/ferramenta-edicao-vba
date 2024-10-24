Attribute VB_Name = "exportarBasketsMultiplas"
Sub exportBasketMultiplas()

    'macro ok
    
'
' EXPORT_BASKET Macro
'
    Application.ScreenUpdating = False
    Dim arqBoletera As Workbook
    Dim boletera As Worksheet
    Dim basket As Worksheet
    Dim export As Worksheet
    Dim dash As Worksheet
    Dim base As Worksheet
'
    Dim endereco As String
    Dim salvar As String
    Dim Data As String
    Dim cliente As String
    Dim broker As String
    Dim testestr
    Dim strpath
    Dim endesave
    Dim salvarBoleta
    Dim nome As String
    Dim testeNA
    Dim arqnome
    Dim wbRece
    Dim Dlin
    Dim ultlin As Integer
    Dim i
    Dim valor
    
    
    Set arqBoletera = ThisWorkbook
    Set boletera = arqBoletera.Sheets("BOLET. ORDENS MÚLTIPLAS")
    Set basket = arqBoletera.Sheets("BASKET - MÚLTIPLAS")
    Set export = arqBoletera.Sheets("EXPORT BSKT MÚLTIPLAS")
    Set dash = arqBoletera.Sheets("DASH BSKT MÚLTIPLAS")
    Set base = arqBoletera.Sheets("BASE")
    
    

    Application.ScreenUpdating = False

    
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
    endesave = CurDir

        
    salvar = endesave & "\3 - RECEITA\" & boletera.Range("H5").Value
    salvarBoleta = endesave & "\2 - BASKETS\"
    modelo = endesave & "\3 - RECEITA\MODELO\RECEITA AVULSA.xlsx"
    If (Dir(salvar, vbDirectory) = "") Then
        MkDir (salvar)
        FileCopy modelo, salvar & "\RECEITA AVULSA.xlsx"
    End If
    salvar = salvar & "\RECEITA AVULSA.xlsx"
    Data = CStr(base.Range("AK8").Value)
    cliente = base.Range("AL7")
    nome = dash.Range("C5").Value
    On Error Resume Next
    broker = base.Range("as6").Value
    codbroker = base.Range("AM6").Value
    
    cont = 1
    For i = 6 To 5000
        ultlin = dash.Cells(i, 3).Row
        valor = dash.Cells(i, 3).Value
        If (valor = "") Then
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
    If (testeNA) Then
        nome = "NOVO"
    End If
    
    arqnome = "(AÇÕES - MÚLTIPLOS) " & Year(Date) & " " & Format(Month(Date), "00") & " " & Format(Day(Date), "00") & " " & nome & " " & boletera.Range("h5").Value & " " & codbroker
    
    export.Range("A1:R500").AutoFilter
    export.Range("A1:R500").AutoFilter
    export.Range("A1:R1", export.Range("A1:R1").End(xlDown)).Copy
    
    Workbooks.Add
    ActiveSheet.Paste
    ActiveWorkbook.SaveAs Filename:= _
        (salvarBoleta & arqnome & ".xlsx") _
        , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWindow.Close
    
    
    'Separando as baskets
    
    For i = 5 To 5000
        valor = dash.Cells(i, 3).Value
        If (valor = "") Then
            Exit For
        Else
            nome = valor
            On Error Resume Next
            testeNA = (boletera.Range("C5").Value = CVErr(xlErrNA))
            If (testeNA) Then
                nome = "NOVO"
            End If
            
            On Error Resume Next
            If dash.Range("D" & i).Value <> "" Then GoTo prox:
            wbRece = Workbooks.Open(salvar, Password:="2022")

            Dlin = Range("A1").End(xlDown).Row + 1
            
            If Dlin > 200000 Then
                Dlin = 2
            End If
            
            Range("A" & Dlin).Value = Data
            Range("B" & Dlin).Value = valor
            Range("C" & Dlin).Value = broker
            
            ActiveWorkbook.Close savechanges:=True
prox:
            arqnome = "(AÇÕES) " & Year(Date) & " " & Format(Month(Date), "00") & " " & Format(Day(Date), "00") & " " & nome & " " & boletera.Range("H5").Value & " " & base.Range("AM6").Value
            
            '--
            export.Range("A1:R500").AutoFilter Field:=1, Criteria1:=nome
            
            '--
            
            export.Range("A1:R1", export.Range("A1:R1").End(xlDown)).Copy
            
            Workbooks.Add
            ActiveSheet.Paste
            ActiveWorkbook.SaveAs Filename:= _
                (salvarBoleta & arqnome & ".xlsx") _
                , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
            ActiveWindow.Close
        End If
    Next

    Application.ScreenUpdating = True
    
    'MsgBox "ok"
    
End Sub

