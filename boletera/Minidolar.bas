Attribute VB_Name = "Minidolar"
Sub minidolar_basket()
    
    codigo = WorksheetFunction.VLookup(ThisWorkbook.Sheets("MINIDOLAR").Range("B10").Value, ThisWorkbook.Sheets("BASE MINIDOLAR").Range("A:B"), 2, 0)
    proxcod = WorksheetFunction.VLookup(ThisWorkbook.Sheets("MINIDOLAR").Range("B10").Value, ThisWorkbook.Sheets("BASE MINIDOLAR").Range("A:D"), 4, 0)
    check = ThisWorkbook.Sheets("MINIDOLAR").Range("J1").Value
    If check = Falso Then
        ticker = "WDO" & codigo & Right(ThisWorkbook.Sheets("MINIDOLAR").Range("C10").Value, 2)
    Else
        If codigo = "Z" Then
            ticker = "WD1" & codigo & Right(ThisWorkbook.Sheets("MINIDOLAR").Range("C10").Value, 2) & proxcod & Right(ThisWorkbook.Sheets("MINIDOLAR").Range("C10").Value + 1, 2)
        Else
            ticker = "WD1" & codigo & Right(ThisWorkbook.Sheets("MINIDOLAR").Range("C10").Value, 2) & proxcod & Right(ThisWorkbook.Sheets("MINIDOLAR").Range("C10").Value, 2)
        End If
    End If
    
    ThisWorkbook.Sheets("BASE MINIDOLAR").Range("F1").Value = "=trade|ult!" & ticker
    ThisWorkbook.Sheets("BASE MINIDOLAR").Range("G1").Value = ticker
    
    ThisWorkbook.Sheets("MINIDOLAR").Range("C16").Value = "TICKER"
    ThisWorkbook.Sheets("MINIDOLAR").Range("D16").Value = "COTAÇÃO"
    ThisWorkbook.Sheets("MINIDOLAR").Range("E16").Value = "SPREAD"
    ThisWorkbook.Sheets("MINIDOLAR").Range("C17").Value = ticker
    ThisWorkbook.Sheets("MINIDOLAR").Range("D17").Value = "='BASE MINIDOLAR'!F1"
    ThisWorkbook.Sheets("MINIDOLAR").Range("C16:E17").Select
    Call formatar
    
    ThisWorkbook.Sheets("MINIDOLAR").Range("G16").Value = "Gerar" & vbNewLine & "Basket"
    ThisWorkbook.Sheets("MINIDOLAR").Range("G16").Select
    Call formatar2
    
End Sub
Sub CaixaDeTexto9_Clique()
    
    Set basket = Workbooks.Add()
    
    If ThisWorkbook.Sheets("MINIDOLAR").Range("J1").Value = Falso Then
        cv = ThisWorkbook.Sheets("MINIDOLAR").Range("H6").Value
    Else
        cv = "Compra"
    End If
    
    With basket.Sheets(1)
        .Range("A1") = "Cliente"
        .Range("B1") = "Qtd."
        .Range("C1") = "Papel"
        .Range("D1") = "Tipo"
        .Range("E1") = "Preço Limite Entrada"
        .Range("F1") = "Preço Disp. Entrada"
        .Range("G1") = "Preço Limite Redução"
        .Range("H1") = "Preço Disp. Redução"
        .Range("I1") = "Preço Limite Objetivo"
        .Range("J1") = "Preço Disp. Objetivo"
        .Range("K1") = "Preço Limite Stop"
        .Range("L1") = "Preço Disp. Stop"
        .Range("M1") = "Preço início"
        .Range("N1") = "Ajuste"
        .Range("O1") = "Validade"
        .Range("P1") = "Dt. Val"
        .Range("Q1") = "Confirmacao"
        .Range("R1") = "Rompimento"
        
        .Range("A2") = ThisWorkbook.Sheets("MINIDOLAR").Range("F6").Value
        .Range("B2") = ThisWorkbook.Sheets("MINIDOLAR").Range("G6").Value
        .Range("C2") = ThisWorkbook.Sheets("BASE MINIDOLAR").Range("G1").Value
        .Range("D2") = cv
        .Range("E2") = ThisWorkbook.Sheets("MINIDOLAR").Range("D17").Value + ThisWorkbook.Sheets("MINIDOLAR").Range("D17").Value * ThisWorkbook.Sheets("MINIDOLAR").Range("E17").Value
        .Range("F2") = 0
        .Range("G2") = 0
        .Range("H2") = 0
        .Range("I2") = 0
        .Range("J2") = 0
        .Range("K2") = 0
        .Range("L2") = 0
        .Range("M2") = 0
        .Range("N2") = 0
        .Range("O2") = "V"
        .Range("P2") = "20130921"
        .Range("Q2") = "1 dia"
        .Range("R2") = ""
    End With
    
End Sub
Sub formatar()

    ThisWorkbook.Sheets("MINIDOLAR").Range("C16:E17").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
End Sub

Sub formatar2()
'
' Macro2 Macro
'

'
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
End Sub
