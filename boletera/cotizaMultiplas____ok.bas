Attribute VB_Name = "cotizaMultiplas____ok"
Option Explicit
Sub cotizaMultiplas()

    Dim arqBoletera As Workbook
    Dim boletera As Worksheet

    Dim lRow As Integer
    Dim tickerAlvo As String
    Dim i As Integer
    Dim funcBullPro As String
    Dim funcBull As String
    Dim funcFinal As String

    'Base/Planilhas
    Set arqBoletera = ThisWorkbook
    Set boletera = arqBoletera.Sheets("BOLET. ORDENS M�LTIPLAS")
    '-----------

    Application.ScreenUpdating = False


    lRow = boletera.Cells(Rows.Count, 4).End(xlUp).Row
    Application.DisplayAlerts = False
    For i = 11 To lRow

    tickerAlvo = boletera.Cells(i, 4)

    If ((Right(boletera.Cells(i, 5).Value, 6) = "COMPRA") Or (Right(boletera.Cells(i, 5).Value, 4) = "GAIN")) Then
        On Error Resume Next
        funcBullPro = "PRODDE|MOFC!" & tickerAlvo
        funcBull = "BULLDDE|MOFC!" & tickerAlvo
        funcFinal = funcBullPro & ";" & funcBull
        funcFinal = "=SEERRO(" & funcFinal & ")"
    Else
        If ((Right(boletera.Cells(i, 5).Value, 5) = "VENDA") Or (Right(boletera.Cells(i, 5).Value, 4) = "LOSS")) Then
            On Error Resume Next
            funcBullPro = "PRODDE|MOFV!" & tickerAlvo
            funcBull = "BULLDDE|MOFV!" & tickerAlvo
            funcFinal = funcBullPro & ";" & funcBull
            funcFinal = "=SEERRO(" & funcFinal & ")"
        Else
            On Error Resume Next
            funcBullPro = "PRODDE|ULT!" & tickerAlvo
            funcBull = "BULLDDE|ULT!" & tickerAlvo
            funcFinal = funcBullPro & ";" & funcBull
            funcFinal = "=SEERRO(" & funcFinal & ")"
        End If
    End If

    If (boletera.Cells(i, 4).Value) = "" Then
        boletera.Cells(i, 32).Value = ""
    Else
        boletera.Cells(i, 31).FormulaLocal = funcFinal
        boletera.Cells(i, 32).Value = tickerAlvo
    End If

    Next

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub




