Attribute VB_Name = "Módulo1"
Sub OrdenaSegunColorRelleno()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rngCol As Range
    Dim colHeader As String
    
    ' Obtener la hoja de trabajo y tabla activa
    Set ws = ActiveSheet
    On Error Resume Next
    Set tbl = ws.ListObjects(ws.Range("A1").ListObject.Name)
    On Error GoTo 0
    
    If tbl Is Nothing Then
        MsgBox "No se pudo identificar la tabla actual. Asegúrate de estar dentro de una tabla.", vbExclamation
        Exit Sub
    End If
    
    ' Solicitar al usuario que seleccione la columna por su encabezado
    colHeader = InputBox("Por favor, introduce el encabezado de la columna (por ejemplo, 'Severidad'):", "Seleccionar columna para ordenar por color")
    
    If colHeader = "" Then Exit Sub ' Si el usuario cancela
    
    ' Verificar si el encabezado existe en la tabla
    On Error Resume Next
    Set rngCol = tbl.ListColumns(colHeader).Range
    On Error GoTo 0
    
    If rngCol Is Nothing Then
        MsgBox "No se encontró el encabezado especificado en la tabla.", vbExclamation
        Exit Sub
    End If
    

     With tbl.Sort
        .SortFields.Clear
        ' Orden rojo (RGB(255, 0, 0))
        .SortFields.Add(rngCol, xlSortOnCellColor, xlAscending, , xlSortNormal).SortOnValue.Color = RGB(255, 255, 0)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' Aplicar el orden por color
    With tbl.Sort
        .SortFields.Clear
        ' Orden amarillo (RGB(0, 176, 80))
        .SortFields.Add(rngCol, xlSortOnCellColor, xlAscending, , xlSortNormal).SortOnValue.Color = RGB(255, 0, 0)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    
    With tbl.Sort
        .SortFields.Clear
        ' Orden morado (RGB(112, 48, 160))
        .SortFields.Add(rngCol, xlSortOnCellColor, xlAscending, , xlSortNormal).SortOnValue.Color = RGB(112, 48, 160)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
   

End Sub

