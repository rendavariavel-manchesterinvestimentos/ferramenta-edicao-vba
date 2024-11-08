Attribute VB_Name = "LongShort"
Sub montaeexporta()
montabasket
exportbasket
End Sub
Sub formulasecotacao()
cotiza
colaformulas
End Sub
Sub montabasketTESTE()
Set atualizador = ThisWorkbook
Set ls = atualizador.Sheets("LONG & SHORT")
Set bskt = atualizador.Sheets("BASKET L&S")
Set export = atualizador.Sheets("EXPORT L&S")

lastrow = ls.Cells(ls.Rows.Count, "B").End(xlUp).Row

linha = 2

For i = 9 To lastrow:
    If ls.Range("B" & i).Value = "" Or ls.Range("D" & i).Value = "" Then GoTo prox
        export.Range("A" & linha).Value = ls.Range("B" & i).Value
        export.Range("B" & linha).Value = ls.Range("E" & i).Value
        export.Range("C" & linha).Value = ls.Range("F" & i).Value
        export.Range("D" & linha).Value = ls.Range("H" & i).Value
        export.Range("E" & linha).Value = ls.Range("K" & i).Value
        export.Range("F" & linha).Value = ls.Range("G" & i).Value
        export.Range("G" & linha).Value = ls.Range("J" & i).Value
        linha = linha + 1
prox:
Next

lastrow = export.Cells(export.Rows.Count, "A").End(xlUp).Row
export.Range("H2").FormulaLocal = "=B2&""F"""
export.Range("I2").FormulaLocal = "=C2&""F"""
export.Range("J2").FormulaLocal = "=SE(CONT.SE(BASE!AO:AP;B2)>0;D2;ARREDMULTB(D2;100))"
export.Range("K2").FormulaLocal = "=SE(CONT.SE(BASE!AO:AP;B2)>0;0;SE(D2<100;D2;D2-J2))"
export.Range("L2").FormulaLocal = "=SE(CONT.SE(BASE!AO:AP;C2)>0;E2;ARREDMULTB(E2;100))"
export.Range("M2").FormulaLocal = "=SE(CONT.SE(BASE!AO:AP;C2)>0;0;SE(E2<100;E2;E2-L2))"

If lastrow > 2 Then
    export.Range("H2").AutoFill Destination:=export.Range("H2:H" & lastrow)
    export.Range("I2").AutoFill Destination:=export.Range("I2:I" & lastrow)
    export.Range("J2").AutoFill Destination:=export.Range("J2:J" & lastrow)
    export.Range("K2").AutoFill Destination:=export.Range("K2:K" & lastrow)
    export.Range("L2").AutoFill Destination:=export.Range("L2:L" & lastrow)
    export.Range("M2").AutoFill Destination:=export.Range("M2:M" & lastrow)
End If

linha = 2

For i = 2 To lastrow:
    If export.Range("J" & i).Value = 0 Then GoTo proxlinha:
        bskt.Range("A" & linha).Value = export.Range("B" & i).Value
        bskt.Range("B" & linha).Value = "COMPRA"
        bskt.Range("C" & linha).Value = export.Range("J" & i).Value
        bskt.Range("D" & linha).Value = export.Range("A" & i).Value
        bskt.Range("E" & linha).Value = export.Range("F" & i).Value
        linha = linha + 1

proxlinha:

    If export.Range("K" & i).Value > 0 Then
        bskt.Range("A" & linha).Value = export.Range("H" & i).Value
        bskt.Range("B" & linha).Value = "COMPRA"
        bskt.Range("C" & linha).Value = export.Range("K" & i).Value
        bskt.Range("D" & linha).Value = export.Range("A" & i).Value
        bskt.Range("E" & linha).Value = export.Range("F" & i).Value
        linha = linha + 1
    End If

    If export.Range("L" & i).Value = 0 Then GoTo proxlinha2:
        bskt.Range("A" & linha).Value = export.Range("C" & i).Value
        bskt.Range("B" & linha).Value = "VENDA"
        bskt.Range("C" & linha).Value = export.Range("L" & i).Value
        bskt.Range("D" & linha).Value = export.Range("A" & i).Value
        bskt.Range("E" & linha).Value = export.Range("G" & i).Value
        linha = linha + 1

proxlinha2:

    If export.Range("M" & i).Value > 0 Then
        bskt.Range("A" & linha).Value = export.Range("I" & i).Value
        bskt.Range("B" & linha).Value = "VENDA"
        bskt.Range("C" & linha).Value = export.Range("M" & i).Value
        bskt.Range("D" & linha).Value = export.Range("A" & i).Value
        bskt.Range("E" & linha).Value = export.Range("G" & i).Value
        linha = linha + 1
    End If

Next

    Call Macro3

End Sub

Sub limpa()
Set atualizador = ThisWorkbook
Set ls = atualizador.Sheets("LONG & SHORT")
Set bskt = atualizador.Sheets("BASKET L&S")
Set export = atualizador.Sheets("EXPORT L&S")

ls.Range("B9", ls.Range("B9").End(xlDown)).ClearContents
ls.Range("D9:V9", ls.Range("D9:V9").End(xlDown)).ClearContents
bskt.Range("A2:E100000").ClearContents
export.Range("A2:M100000").ClearContents

End Sub

Sub cotiza()

    Set atualizador = ThisWorkbook
    Set ls = atualizador.Sheets("LONG & SHORT")

    Application.ScreenUpdating = False

    lastrow = ls.Cells(Rows.Count, 5).End(xlUp).Row
    Application.DisplayAlerts = False
    For i = 9 To lastrow
        tickercompra = ls.Cells(i, 5)
        tickervenda = ls.Cells(i, 6)

        cfuncBull = "BULLDDE|MOFC!" & tickercompra
        cfuncBullPro = "PRODDE|MOFC!" & tickercompra
        cfuncFinal = cfuncBull & ";" & cfuncBullPro
        cfuncFinal = "=SEERRO(" & cfuncFinal & ")*(1+$F$2)"

        vfuncBullPro = "PRODDE|MOFV!" & tickervenda
        vfuncBull = "BULLDDE|MOFV!" & tickervenda
        vfuncFinal = vfuncBullPro & ";" & vfuncBull
        vfuncFinal = "=SEERRO(" & vfuncFinal & ")*(1-$F$3)"

        ls.Cells(i, 7).FormulaLocal = cfuncFinal
        ls.Cells(i, 10).FormulaLocal = vfuncFinal

    Next

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub
Sub validacao()


Set atualizador = ThisWorkbook
Set ls = atualizador.Sheets("LONG & SHORT")
Set export = atualizador.Sheets("BASKET L&S")

ThisWorkbook.Unprotect Password:="senhadaboletera"
ThisWorkbook.Sheets("BASKET L&S").Visible = True

If IsEmpty(ls.Range("C3")) Then
MsgBox ("Coloque o seu código de broker na célula C3")
Else
Call montabasketTESTE
Call exportbasket
Call receita

ThisWorkbook.Protect Structure:=True, Windows:=False, Password:="senhadaboletera"

End If

End Sub
Sub exportbasket()

    ' variável global
    EstaPastaDeTrabalho.Importar_Variaveis_Globais

    Set atualizador = ThisWorkbook
    Set ls = atualizador.Sheets("LONG & SHORT")
    Set export = atualizador.Sheets("BASKET L&S")

    ' caminho usando ONEDRIVE_GERAL
    Dim caminhoBasket As String
    caminhoBasket = fso.BuildPath(ONEDRIVE_GERAL, "Ferramentas\Boletera\Baskets\Long e short")

    Application.ScreenUpdating = False

    broker = ls.Range("C3").Value
    valor = ""

    For i = 2 To 5000

        If valor = export.Cells(i, 4).Value Then GoTo oi

        valor = export.Cells(i, 4).Value
        If (valor = "") Then
            Exit For
        Else
            nome = valor

            arqnome = "(L&S) " & Year(Date) & " " & Format(Month(Date), "00") & " " & Format(Day(Date), "00") & " " & nome & " " & broker

            export.Range("A1:E100000").AutoFilter Field:=4, Criteria1:=nome

            export.Range("A1:E1", export.Range("A1:E1").End(xlDown)).Copy

            Workbooks.Add
            ActiveSheet.Paste
            ActiveWorkbook.Sheets.Add

            Set basket = ActiveWorkbook.Sheets("Planilha1")
            Set basketemail = ActiveWorkbook.Sheets("Planilha2")
            basketemail.Name = "Auditoria Email"
            basket.Move before:=Sheets(1)

            basketemail.Range("A5:D5").Merge
            basketemail.Range("A5") = "ORDENS A MERCADO"
            basketemail.Range("A6") = "Cliente"
            basketemail.Range("B6") = "Ativo"
            basketemail.Range("C6") = "C/V"
            basketemail.Range("D6") = "Qtd. Total"

            lastrow = 5 + ActiveWorkbook.Sheets("Planilha1").Cells(Rows.Count, 5).End(xlUp).Row

            basketemail.Range("A7").FormulaLocal = "=SE('Planilha1'!D2="""";"""";'Planilha1'!D2)"
            basketemail.Range("B7").FormulaLocal = "=SE('Planilha1'!A2="""";"""";'Planilha1'!A2)"
            basketemail.Range("C7").FormulaLocal = "=SE('Planilha1'!B2="""";"""";'Planilha1'!B2)"
            basketemail.Range("D7").FormulaLocal = "=SE('Planilha1'!C2="""";"""";'Planilha1'!C2)"

            basketemail.Range("A7").AutoFill Destination:=basketemail.Range("A7:A" & lastrow)
            basketemail.Range("B7").AutoFill Destination:=basketemail.Range("B7:B" & lastrow)
            basketemail.Range("C7").AutoFill Destination:=basketemail.Range("C7:C" & lastrow)
            basketemail.Range("D7").AutoFill Destination:=basketemail.Range("D7:D" & lastrow)

            basketemail.Range("A5:D" & lastrow).HorizontalAlignment = xlCenter
            basketemail.Range("A5:D" & lastrow).VerticalAlignment = xlCenter

            basketemail.Range("A5:D" & lastrow).Borders(xlDiagonalDown).LineStyle = xlNone
            basketemail.Range("A5:D" & lastrow).Borders(xlDiagonalUp).LineStyle = xlNone

            With basketemail.Range("A5:D" & lastrow).Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With basketemail.Range("A5:D" & lastrow).Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With basketemail.Range("A5:D" & lastrow).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With basketemail.Range("A5:D" & lastrow).Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With basketemail.Range("A5:D" & lastrow).Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With basketemail.Range("A5:D" & lastrow).Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            basketemail.Range("A5:D6").Font.Bold = True
            basketemail.Range("A1") = "Prezado(a),"
            basketemail.Range("A3") = "Você autoriza todas as operações descritas abaixo?"

            ActiveWorkbook.SaveAs Filename:=fso.BuildPath(caminhoBasket, arqnome & ".xlsx"), FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
            ActiveWindow.Close

oi:
        End If
    Next

    export.Range("A1:E100000").AutoFilter Field:=4

    Application.ScreenUpdating = True

End Sub


Sub colaformulas()

Set atualizador = ThisWorkbook
Set ls = atualizador.Sheets("LONG & SHORT")

Application.ScreenUpdating = False

lastrow = ls.Cells(Rows.Count, 2).End(xlUp).Row

ls.Range("H9").FormulaLocal = "=SE(G9="""";"""";SEERRO(ARRED(((1-$F$4)*D9*(1+$F$2))/G9;0);""Sem preço""))"
ls.Range("I9").FormulaLocal = "=G9*H9"

ls.Range("K9").FormulaLocal = "=SE(I9="""";"""";SEERRO(ARRED((D9*(1-$F$3)/J9);0);""Sem preço""))"
ls.Range("L9").FormulaLocal = "=J9*K9"
ls.Range("M9").FormulaLocal = "=L9-I9"
ls.Range("N9").FormulaLocal = "=SE(D9="""";"""";SE(CONT.SE($B$9:B9;B9)>1;(I9+L9)*0,5%;(I9+L9)*0,5%+25,21))"
ls.Range("O9").FormulaLocal = "=M9-N9"
ls.Range("P9").FormulaLocal = "=SE(B9="""";"""";SE(CONT.SE($B$9:B9;B9)>1;""Linha "" & PROCV(B9;B:X;22;0);SEERRO(PROCV(B9;BASE!N:Q;4;0);""Fora da base"")))"
ls.Range("Q9").FormulaLocal = "=SE(B9="""";"""";SE(CONT.SE($B$9:B9;B9)>1;""Linha "" & PROCV(B9;B:X;22;0); P9+SOMASE(B:F;B9;O:O)))"
ls.Range("R9").FormulaLocal = "=SEERRO(G9/J9;""Verificar"")"
ls.Range("S9").FormulaLocal = "=SE(D9="""";"""";SE(CONT.SE($B$9:B9;B9)>1;""Linha "" & PROCV(B9;B:X;22;0);2*SOMASE(B:B;B9;D:D)))"

ls.Range("T9").FormulaLocal = "=SE(D9="""";"""";SE(CONT.SE($B$9:B9;B9)>1;""Linha "" & PROCV(B9;B:X;22;0);SEERRO(PROCV(B9;Garantia!L:M;2;0);""Fora da base"")))"
ls.Range("U9").FormulaLocal = "=SE(ESQUERDA(S9;3)=""Ver"";""Linha "" & PROCV(B9;B:X;22;0);SEERRO(S9/PROCV(B9;Alocação!A:C;3;0);""Fora da base""))"
ls.Range("V9").FormulaLocal = "=SE(OU(E9="""";F9="""");"""";SE(CONT.SES(BASE!E:E;B9;BASE!F:F;E9)+CONT.SES(BASE!E:E;B9;BASE!F:F;F9);""Verificar custódia"";""OK""))"

If lastrow > 9 Then
    ls.Range("H9").AutoFill Destination:=ls.Range("H9:H" & lastrow)
    ls.Range("I9").AutoFill Destination:=ls.Range("I9:I" & lastrow)
    ls.Range("K9").AutoFill Destination:=ls.Range("K9:K" & lastrow)
    ls.Range("L9").AutoFill Destination:=ls.Range("L9:L" & lastrow)
    ls.Range("M9").AutoFill Destination:=ls.Range("M9:M" & lastrow)
    ls.Range("N9").AutoFill Destination:=ls.Range("N9:N" & lastrow)
    ls.Range("O9").AutoFill Destination:=ls.Range("O9:O" & lastrow)
    ls.Range("P9").AutoFill Destination:=ls.Range("P9:P" & lastrow)
    ls.Range("Q9").AutoFill Destination:=ls.Range("Q9:Q" & lastrow)
    ls.Range("R9").AutoFill Destination:=ls.Range("R9:R" & lastrow)
    ls.Range("S9").AutoFill Destination:=ls.Range("S9:S" & lastrow)
    ls.Range("T9").AutoFill Destination:=ls.Range("T9:T" & lastrow)
    ls.Range("U9").AutoFill Destination:=ls.Range("U9:U" & lastrow)
    ls.Range("V9").AutoFill Destination:=ls.Range("V9:V" & lastrow)
End If

End Sub
Sub receita()

Application.ScreenUpdating = False

' variável global
EstaPastaDeTrabalho.Importar_Variaveis_Globais

Set atualizador = ThisWorkbook
Set ls = atualizador.Sheets("LONG & SHORT")
Set base = atualizador.Sheets("BASE")
Set export = atualizador.Sheets("BASKET L&S")

data = base.Range("AK8").Value

broker = ls.Range("C4").Value
codigo = ls.Range("C3").Value

lastrow = ls.Cells(Rows.Count, 2).End(xlUp).Row

' caminho usando a variável global
Dim caminhoReceita As String
caminhoReceita = fso.BuildPath(ONEDRIVE_GERAL, "Ferramentas\Boletera\Receita\" & broker)

' abre o arquivo com o caminho modificado
Set arquivobroker = Workbooks.Open(caminhoReceita & "\RECEITA AVULSA.xlsx", Password:="2022")
Set abaplanilha = arquivobroker.Sheets("Plan1")

lastrowreceita = abaplanilha.Cells(Rows.Count, 2).End(xlUp).Row + 1
ls.Range("B9:B" & lastrow).Copy
abaplanilha.Range("B" & lastrowreceita).PasteSpecial xlPasteValues

lastrowreceita2 = abaplanilha.Cells(Rows.Count, 2).End(xlUp).Row

abaplanilha.Range("A" & lastrowreceita & ":A" & lastrowreceita2).Value = data
abaplanilha.Range("C" & lastrowreceita & ":C" & lastrowreceita2).Value = codigo

arquivobroker.Close savechanges:=True
export.Activate

Application.ScreenUpdating = True

End Sub

Sub Macro3()

    ActiveWorkbook.Worksheets("BASKET L&S").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("BASKET L&S").AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("D1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("BASKET L&S").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub


