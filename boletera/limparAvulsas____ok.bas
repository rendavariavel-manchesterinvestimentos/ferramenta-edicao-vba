Attribute VB_Name = "limparAvulsas____ok"
Option Explicit
Sub LIMPAR_CUST�DIA()
'
' LIMPAR_CUST�DIA Macro
'
    Dim arqBoletera As Workbook
    Dim base As Worksheet

    Set arqBoletera = ThisWorkbook
    Set base = arqBoletera.Sheets("BASE")

    base.Range("E7:H100000").ClearContents

End Sub

Sub LIMPAR_SALDOS()
'
' LIMPAR_SALDOS Macro
'
    Dim arqBoletera As Workbook
    Dim base As Worksheet

    Set arqBoletera = ThisWorkbook
    Set base = arqBoletera.Sheets("BASE")


    base.Range("N7:U10000").Select

End Sub

