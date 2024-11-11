Attribute VB_Name = "copiarQtdAtual__ok"
Option Explicit
Sub copiarQtdAtualAvulsas()
'
' COPIAR_QTD_ATUAL Macro
'
    Dim arqBoletera As Workbook
    Dim boletera As Worksheet

    Set arqBoletera = ThisWorkbook
    Set boletera = arqBoletera.Sheets("BOLET. AVULSAS")

    boletera.Range("f11:f80").Copy
    boletera.Range("K11").PasteSpecial xlPasteValues
End Sub

Sub copiarQtdAtualMultiplas()
'
' COPIAR_QTD_ATUAL Macro
'
    Dim arqBoletera As Workbook
    Dim boletera As Worksheet

    Set arqBoletera = ThisWorkbook
    Set boletera = arqBoletera.Sheets("BOLET. ORDENS M�LTIPLAS")

    boletera.Range("H11:H80").Copy
    boletera.Range("m11").PasteSpecial xlPasteValues
End Sub


