Attribute VB_Name = "Módulo1"
Sub Teste()
Attribute Teste.VB_ProcData.VB_Invoke_Func = " \n14"

Set boletera = ThisWorkbook
Set boleta = boletera.Sheets("BOLET. AVULSAS")
Set base = boletera.Sheets("BASE")

    fim = base.Range("AU7").End(xlDown).Row
    cliente = WorksheetFunction.CountIf(Range("AU7:AU" & fim), boleta.Range("c4").Value)
    
    If cliente = 0 Then GoTo basket
    Else
    ultlinha = boleta.Range("A10").End(xlDown).Row
    For i = 11 To ultlinha
        If Worksheet.Function.CountIf(Range("AV7:AV16"), boleta.Range("A" & i).Value) > 0 Then
        MsgBox "O cliente está na Dinâmica, NÃO OPERE ESSE ATIVO"
        End
        End If
    Next
    End If
    
basket:
    
End Sub
