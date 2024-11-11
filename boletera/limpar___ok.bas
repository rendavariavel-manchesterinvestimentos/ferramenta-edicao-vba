Attribute VB_Name = "limpar___ok"
Option Explicit
Sub limparMultiplas()
Attribute limparMultiplas.VB_ProcData.VB_Invoke_Func = " \n14"

    Dim arqBoletera As Workbook
    Dim boleta As Worksheet

    Set arqBoletera = ThisWorkbook
    Set boleta = arqBoletera.Sheets("BOLET. ORDENS M�LTIPLAS")

    With boleta
        .Range("M11:M80").ClearContents
        .Range("B11:b80").ClearContents
        .Range("d11:e80").ClearContents
        .Range("AE11:AF80").ClearContents
    End With

    'MsgBox ("Feito")

End Sub
Sub limparAvulsas()

    Dim arqBoletera As Workbook
    Dim boleta As Worksheet

    Set arqBoletera = ThisWorkbook
    Set boleta = arqBoletera.Sheets("BOLET. AVULSAS")

    With boleta
        .Range("C4").ClearContents
        .Range("B11:D80").ClearContents
        .Range("K11:K80").ClearContents
        .Range("AE11:AF80").ClearContents
    End With
    'MsgBox ("Feito")

End Sub

