Attribute VB_Name = "Módulo1"
Sub MarkInlineCharts()
    Dim doc As Document
    Dim inlineShape As inlineShape
    Dim count As Integer
    count = 1
    
    Set doc = ActiveDocument ' Documento activo
    
    For Each inlineShape In doc.InlineShapes
        On Error Resume Next ' Ignorar errores para elementos sin texto alternativo
        inlineShape.AlternativeText = "MyGrafico " & CStr(count)
        On Error GoTo 0 ' Restaurar manejo normal de errores
        count = count + 1
    Next inlineShape
End Sub

