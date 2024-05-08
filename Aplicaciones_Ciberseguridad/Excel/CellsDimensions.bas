Attribute VB_Name = "Módulo1"
Sub FormatoCeldas()

    Dim rng As Range
    Dim cell As Range
    
    ' Verificar si se ha seleccionado un rango de celdas
    If TypeName(Selection) <> "Range" Then
        MsgBox "Por favor, selecciona un rango de celdas primero.", vbExclamation
        Exit Sub
    End If
    
    Set rng = Selection
    
    ' Establecer altura de fila y ancho de celda
    rng.EntireRow.RowHeight = 15
    rng.EntireColumn.AutoFit
    
    ' Verificar el ancho de las columnas
    For Each cell In rng
        If cell.ColumnWidth > 32 Then
            cell.EntireColumn.ColumnWidth = 32
        End If
    Next cell

End Sub

