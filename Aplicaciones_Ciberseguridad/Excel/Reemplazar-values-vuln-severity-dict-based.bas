Attribute VB_Name = "Módulo1"
Sub OrdenaSegunColorRelleno()
Attribute OrdenaSegunColorRelleno.VB_ProcData.VB_Invoke_Func = " \n14"
'
' OrdenaSegunColorRelleno Macro
'

'
    ActiveWorkbook.Worksheets("Vulnerabilidades").ListObjects("Tabla_vulnerabilidades").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Vulnerabilidades").ListObjects("Tabla_vulnerabilidades").Sort. _
        SortFields.Add(Range("Tabla_vulnerabilidades[[#All],[Severidad]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(0, 176, 80)
    With ActiveWorkbook.Worksheets("Vulnerabilidades").ListObjects("Tabla_vulnerabilidades").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Vulnerabilidades").ListObjects("Tabla_vulnerabilidades").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Vulnerabilidades").ListObjects("Tabla_vulnerabilidades").Sort. _
        SortFields.Add(Range("Tabla_vulnerabilidades[[#All],[Severidad]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(255, 255, 0)
    With ActiveWorkbook.Worksheets("Vulnerabilidades").ListObjects("Tabla_vulnerabilidades").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Vulnerabilidades").ListObjects("Tabla_vulnerabilidades").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Vulnerabilidades").ListObjects("Tabla_vulnerabilidades").Sort. _
        SortFields.Add(Range("Tabla_vulnerabilidades[[#All],[Severidad]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(255, 0, 0)
    With ActiveWorkbook.Worksheets("Vulnerabilidades").ListObjects("Tabla_vulnerabilidades").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
