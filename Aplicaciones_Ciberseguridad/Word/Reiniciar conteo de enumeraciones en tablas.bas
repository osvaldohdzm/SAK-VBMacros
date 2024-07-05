Attribute VB_Name = "Módulo32"
Sub ReiniciarConteoEnumeracionesEnTablas()
    Dim tabla As Table
    Dim rngCell As Range
    Dim subcadena As String
    Dim estiloLista As String
    Dim i As Integer
    
    ' Definir la subcadena deseada y el estilo de lista deseado
    subcadena = "REFERENCIAS"
    estiloLista = "Viñeta referencia"
    
    ' Iterar sobre todas las tablas en el documento
    For Each tabla In ActiveDocument.Tables
        ' Verificar si la tabla tiene al menos dos columnas
        If tabla.Columns.Count >= 2 Then
            ' Iterar sobre todas las filas de la tabla
            For i = 1 To tabla.Rows.Count
                ' Obtener el contenido de la celda en la primera columna de la fila actual
                Set rngCell = tabla.Cell(i, 1).Range
                rngCell.End = rngCell.End - 1 ' Eliminar el salto de línea al final de la celda
                
                ' Verificar si la celda contiene la subcadena deseada
                If ContieneSubcadena(rngCell.Text, subcadena) Then
                    ' Obtener la celda en la segunda columna de la fila actual
                    Set rngCell = tabla.Cell(i, 2).Range
                    rngCell.End = rngCell.End - 1 ' Eliminar el salto de línea al final de la celda
                    
                    ' Aplicar el reinicio del conteo numérico
                    rngCell.ListFormat.ApplyListTemplateWithLevel ListTemplate:= _
                        ListGalleries(wdNumberGallery).ListTemplates(1), ContinuePreviousList:= _
                        False, ApplyTo:=wdListApplyToWholeList
                        
                    Exit For ' Salir del bucle una vez que se reinicia la enumeración
                End If
            Next i
        End If
    Next tabla
End Sub

Function ContieneSubcadena(ByVal texto As String, ByVal subcadena As String) As Boolean
    ContieneSubcadena = InStr(1, texto, subcadena, vbTextCompare) > 0
End Function


