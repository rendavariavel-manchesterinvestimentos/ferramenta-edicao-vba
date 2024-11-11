Attribute VB_Name = "Carteiras"
Sub copiar_dinamica()

    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("B11:B80").ClearContents
    ThisWorkbook.Sheets("BASE").Range("AA7:AA19").Copy
    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("B11").PasteSpecial xlPasteValues
    
    linha = ThisWorkbook.Sheets("BOLET. AVULSAS").Range("B100000").End(xlUp).Row
    
    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("C11:C80").ClearContents
    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("C11:C" & linha).Value = "COMPRA"
    
    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("D11:D80").ClearContents
    ThisWorkbook.Sheets("BASE").Range("AB7:AB19").Copy
    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("D11").PasteSpecial xlPasteValues
    
End Sub
Sub copiar_dinamica_RENDA()

    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("B11:B80").ClearContents
    ThisWorkbook.Sheets("BASE").Range("AA23:AA37").Copy
    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("B11").PasteSpecial xlPasteValues
    
    linha = ThisWorkbook.Sheets("BOLET. AVULSAS").Range("B100000").End(xlUp).Row
    
    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("C11:C80").ClearContents
    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("C11:C" & linha).Value = "COMPRA"
    
    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("D11:D80").ClearContents
    ThisWorkbook.Sheets("BASE").Range("AB23:AB37").Copy
    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("D11").PasteSpecial xlPasteValues
    
End Sub
Sub copiar_dividendos()

    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("B11:B80").ClearContents
    ThisWorkbook.Sheets("BASE").Range("AA41:AA50").Copy
    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("B11").PasteSpecial xlPasteValues
    
    linha = ThisWorkbook.Sheets("BOLET. AVULSAS").Range("B100000").End(xlUp).Row
    
    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("C11:C80").ClearContents
    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("C11:C" & linha).Value = "COMPRA"
    
    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("D11:D80").ClearContents
    ThisWorkbook.Sheets("BASE").Range("AB41:AB50").Copy
    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("D11").PasteSpecial xlPasteValues

End Sub

Sub copiar_dinamica_FII()

    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("B11:B80").ClearContents
    ThisWorkbook.Sheets("BASE").Range("AD7:AD30").Copy
    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("B11").PasteSpecial xlPasteValues
    
    linha = ThisWorkbook.Sheets("BOLET. AVULSAS").Range("B100000").End(xlUp).Row
    
    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("C11:C80").ClearContents
    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("C11:C" & linha).Value = "COMPRA"
    
    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("D11:D80").ClearContents
    ThisWorkbook.Sheets("BASE").Range("AE7:AE30").Copy
    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("D11").PasteSpecial xlPasteValues
    
End Sub
Sub troca_Carteiras()

     ThisWorkbook.Sheets("BOLET. AVULSAS").Range("B11:B80").ClearContents
    ThisWorkbook.Sheets("BASE").Range("AD30:AD43").Copy
    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("B11").PasteSpecial xlPasteValues
    
    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("C11:C80").ClearContents
    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("C11:C18").Value = "VENDA"
    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("C19:C24").Value = "COMPRA"
    
    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("D11:D80").ClearContents
    ThisWorkbook.Sheets("BASE").Range("AE30:AE43").Copy
    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("D11").PasteSpecial xlPasteValues

    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("I5").FormulaLocal = "=seerro(SOMA(N11:N17)*0,995-25,21;0)"
    
    With ThisWorkbook.Sheets("BOLET. AVULSAS")
        .Range("K11").FormulaLocal = "=F11"
        .Range("K12").FormulaLocal = "=F12"
        .Range("K13").FormulaLocal = "=F13"
        .Range("K14").FormulaLocal = "=F14"
        .Range("K15").FormulaLocal = "=F15"
        .Range("K16").FormulaLocal = "=F16"
        .Range("K17").FormulaLocal = "=F17"
        .Range("K18").FormulaLocal = "=F18"
        .Range("K19").FormulaLocal = "=I19"
        .Range("K20").FormulaLocal = "=I20"
        .Range("K21").FormulaLocal = "=I21"
        .Range("K22").FormulaLocal = "=I22"
        .Range("K23").FormulaLocal = "=I23"
        .Range("K24").FormulaLocal = "=I24"
        .Range("K25").FormulaLocal = "=I25"
        .Range("K26").FormulaLocal = "=I26"
        
        
    End With

    
End Sub
Sub copiar_smallcaps()

    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("B11:B80").ClearContents
    ThisWorkbook.Sheets("BASE").Range("AA52:AA61").Copy
    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("B11").PasteSpecial xlPasteValues
    
    linha = ThisWorkbook.Sheets("BOLET. AVULSAS").Range("B100000").End(xlUp).Row
    
    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("C11:C80").ClearContents
    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("C11:C" & linha).Value = "COMPRA"
    
    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("D11:D80").ClearContents
    ThisWorkbook.Sheets("BASE").Range("AB52:AB61").Copy
    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("D11").PasteSpecial xlPasteValues

End Sub

Sub copiar_WISIRALUGUELDEIMÓVEIS()

    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("B11:B80").ClearContents
    ThisWorkbook.Sheets("BASE").Range("AA66:AA77").Copy
    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("B11").PasteSpecial xlPasteValues
    
    linha = ThisWorkbook.Sheets("BOLET. AVULSAS").Range("B100000").End(xlUp).Row
    
    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("C11:C80").ClearContents
    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("C11:C" & linha).Value = "COMPRA"
    
    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("D11:D80").ClearContents
    ThisWorkbook.Sheets("BASE").Range("AB66:AB77").Copy
    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("D11").PasteSpecial xlPasteValues

End Sub
Sub copiar_WISIRFIISSTARTER()

    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("B11:B80").ClearContents
    ThisWorkbook.Sheets("BASE").Range("AA82:AA93").Copy
    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("B11").PasteSpecial xlPasteValues
    
    linha = ThisWorkbook.Sheets("BOLET. AVULSAS").Range("B100000").End(xlUp).Row
    
    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("C11:C80").ClearContents
    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("C11:C" & linha).Value = "COMPRA"
    
    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("D11:D80").ClearContents
    ThisWorkbook.Sheets("BASE").Range("AB82:AB93").Copy
    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("D11").PasteSpecial xlPasteValues

End Sub

Sub troca_rendaparadinamica()

 ThisWorkbook.Sheets("BOLET. AVULSAS").Range("B11:B80").ClearContents
    ThisWorkbook.Sheets("BASE").Range("AD47:AD65").Copy
    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("B11").PasteSpecial xlPasteValues
    
    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("C11:C80").ClearContents
    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("C11:C16").Value = "VENDA"
    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("C17:C27").Value = "COMPRA"
    
    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("D11:D80").ClearContents
    ThisWorkbook.Sheets("BASE").Range("AE47:AE63").Copy
    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("D11").PasteSpecial xlPasteValues

    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("I5").FormulaLocal = "=seerro(SOMA(N11:N17)*0,995-25,21;0)"
    
    With ThisWorkbook.Sheets("BOLET. AVULSAS")
        .Range("K11").FormulaLocal = "=F11"
        .Range("K12").FormulaLocal = "=F12"
        .Range("K13").FormulaLocal = "=F13"
        .Range("K14").FormulaLocal = "=F14"
        .Range("K15").FormulaLocal = "=F15"
        .Range("K16").FormulaLocal = "=F16"
        .Range("K17").FormulaLocal = "=i17"
        .Range("K18").FormulaLocal = "=i18"
        .Range("K19").FormulaLocal = "=I19"
        .Range("K20").FormulaLocal = "=I20"
        .Range("K21").FormulaLocal = "=I21"
        .Range("K22").FormulaLocal = "=I22"
        .Range("K23").FormulaLocal = "=I23"
        .Range("K24").FormulaLocal = "=I24"
        .Range("K25").FormulaLocal = "=I25"
        .Range("K26").FormulaLocal = "=I26"
        .Range("K27").FormulaLocal = "=I27"
        .Range("K28").FormulaLocal = "=I28"
        .Range("K29").FormulaLocal = "=I29"
        
    End With

End Sub

Sub troca_Carteiras_ALT()

    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("B11:B80").ClearContents
    ThisWorkbook.Sheets("BASE").Range("AD30:AD37").Copy
    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("B11").PasteSpecial xlPasteValues
    
    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("C11:C80").ClearContents
    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("C11:C14").Value = "VENDA"
    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("C15:C18").Value = "COMPRA"
    
    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("D11:D80").ClearContents
    ThisWorkbook.Sheets("BASE").Range("AE30:AE37").Copy
    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("D11").PasteSpecial xlPasteValues

    ThisWorkbook.Sheets("BOLET. AVULSAS").Range("I5").FormulaLocal = "=seerro(SOMA(N11:N14)*0,995-25,21;0)"
    
    With ThisWorkbook.Sheets("BOLET. AVULSAS")
        .Range("K11").FormulaLocal = "=F11"
        .Range("K12").FormulaLocal = "=F12"
        .Range("K13").FormulaLocal = "=F13"
        .Range("K14").FormulaLocal = "=F14"
        .Range("K15").FormulaLocal = "=I15"
        .Range("K16").FormulaLocal = "=I16"
        .Range("K17").FormulaLocal = "=I17"
        .Range("K18").FormulaLocal = "=I18"
    End With

    
End Sub
Sub aviso()
 MsgBox ("Não é pra clicar pora")
End Sub

Sub gerar_arquivo()

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    EstaPastaDeTrabalho.Importar_Variaveis_Globais

    Application.ScreenUpdating = False
    ThisWorkbook.Unprotect "senhadaboletera"
    ThisWorkbook.Worksheets("MODELO TOMBAMENTO").Visible = True
    ThisWorkbook.Worksheets("MODELO TOMBAMENTO").Copy
    ThisWorkbook.Worksheets("MODELO TOMBAMENTO").Visible = False
    ThisWorkbook.Protect "senhadaboletera"
    Application.ScreenUpdating = True
    
    ' Import da variável global
    EstaPastaDeTrabalho.Importar_Variaveis_Globais
    
    Set arquivinho = Application.ActiveSheet
    linha = arquivinho.Range("B10000").End(xlUp).Row
    arquivinho.Range("B3:D" & linha).Copy
    arquivinho.Range("B3:D" & linha).PasteSpecial xlPasteValues
    nomearq = "MÁSCARA CARTEIRA " & arquivinho.Range("B3").Value
    
    Dim caminhoHub As String
    caminhoHub = fso.BuildPath(EstaPastaDeTrabalho.ONEDRIVE_GERAL, "Ferramentas\Boletera\Carteiras")
    
    ActiveWorkbook.SaveAs Filename:=fso.BuildPath(caminhoHub, nomearq & ".xlsx")
    
End Sub