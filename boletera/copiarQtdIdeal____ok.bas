Attribute VB_Name = "copiarQtdIdeal____ok"
Option Explicit
Sub copiarQtdIdealAvulsas()
Attribute copiarQtdIdealAvulsas.VB_ProcData.VB_Invoke_Func = " \n14"
'
' COPIAR_QTD_IDEAL Macro
'
    Dim arqBoletera As Workbook
    Dim boletera As Worksheet
    
    Set arqBoletera = ThisWorkbook
    Set boletera = arqBoletera.Sheets("BOLET. AVULSAS")
    
    Range("I11:I80").Copy
    Range("K11").PasteSpecial xlPasteValues
End Sub
Sub copiarQtdIdealMultiplas()

    Dim arqBoletera As Workbook
    Dim boletera As Worksheet
    
    Set arqBoletera = ThisWorkbook
    Set boletera = arqBoletera.Sheets("BOLET. ORDENS MÚLTIPLAS")
    
    Range("k11:k80").Copy
    Range("m11").PasteSpecial xlPasteValues
    
End Sub
