Attribute VB_Name = "basketMultiClientes"
Option Explicit
Sub basketMultiClientes()

    Dim boletera            As Worksheet
    Dim basket              As Worksheet
    Dim basketForm          As Worksheet
    Dim dashBasket          As Worksheet
    Dim arquivoBoletera     As Workbook
    
    Dim ultlin As Integer
    Dim i
    Dim valor
    
        
    'Def variaveis
        'Workbooks
    Set arquivoBoletera = ThisWorkbook
    
        'Worksheets
    Set boletera = arquivoBoletera.Worksheets("BOLET. ORDENS MÚLTIPLAS")
    Set basket = arquivoBoletera.Worksheets("EXPORT BSKT MÚLTIPLAS")
    Set basketForm = arquivoBoletera.Worksheets("BASKET - MÚLTIPLAS")
    Set dashBasket = arquivoBoletera.Worksheets("DASH BSKT MÚLTIPLAS")
    '----------------------
    
    Application.ScreenUpdating = False
    
    
    'check se B11 tá vazio
    If IsEmpty(boletera.Range("B11")) Then
        MsgBox ("O começo da lista (B11) está vazio")
        Exit Sub
    End If
    '-------------------
    
    
    
    basket.Range("A1:R500").AutoFilter
    basket.Range("A1:R500").AutoFilter
    basketForm.Range("A3:R200").Copy
    basket.Range("A2").PasteSpecial xlPasteValues
    
    With basket.AutoFilter.Sort
        .SortFields.Clear
        .SortFields.Add2 key _
            :=Range("A1:A10000"), SortOn:=xlSortOnValues, Order:=xlAscending, _
            DataOption:=xlSortNormal
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    
    For i = 2 To 5000
        ultlin = basket.Cells(i, 1).Row
        valor = basket.Cells(i, 1).Value
        If (valor = "") Then
            Exit For
        End If
    Next
    ultlin = ultlin - 1
    
    basket.Range("A2:A" & ultlin).Copy
    
    
    With dashBasket
        .Range("C5").PasteSpecial xlPasteValues
        .Range("$C4:$C1000").RemoveDuplicates Columns:=1, Header:=xlYes
    End With
    '-------------------
    
    Application.ScreenUpdating = True
    
    dashBasket.Select
End Sub

Sub limparMultiClientes()

    Application.ScreenUpdating = False
    Dim dashBasket          As Worksheet
    Dim arquivoBoletera     As Workbook
    Dim export As Worksheet
    
     
    Set arquivoBoletera = ThisWorkbook
    
    'Worksheets
    Set dashBasket = arquivoBoletera.Worksheets("DASH BSKT MÚLTIPLAS")
    Set export = arquivoBoletera.Worksheets("EXPORT BSKT MÚLTIPLAS")
    '----------------------

    'limpar células
    With dashBasket
        .Range(.Range("C5"), .Range("C5").End(xlDown)).ClearContents
    End With
    
    export.Range("A2:R200").ClearContents
    
    Application.ScreenUpdating = True
    'MsgBox ("Basket limpa")
End Sub
