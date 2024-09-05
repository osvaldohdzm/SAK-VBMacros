Attribute VB_Name = "Módulo1"
Sub MarkTables()
    Dim doc As Document
    Dim table As table
    Dim count As Integer
    count = 1
    
    Set doc = ActiveDocument ' Documento activo
    
    For Each table In doc.Tables
        On Error Resume Next ' Ignorar errores para tablas sin celdas
        table.Cell(1, 1).Range.Text = "Table " & CStr(count)
        On Error GoTo 0 ' Restaurar manejo normal de errores
        count = count + 1
    Next table
End Sub
