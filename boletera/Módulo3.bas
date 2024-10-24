Attribute VB_Name = "Módulo3"
Sub Macro3()
Attribute Macro3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro3 Macro
'

'
    ActiveWorkbook.Worksheets("BASKET L&S").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("BASKET L&S").AutoFilter.Sort.SortFields.Add2 key:= _
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
