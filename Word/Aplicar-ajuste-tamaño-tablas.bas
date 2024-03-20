Attribute VB_Name = "Módulo1"
Sub AjustaTamanoTablasInternasAnidads()
    Dim tabla As Table
    Dim celda As cell
    Dim tablaInterna As Table

    ' Recorrer todas las tablas en el documento
    For Each tabla In ActiveDocument.Tables

        ' Recorrer todas las celdas de la tabla para buscar tablas internas
        For Each celda In tabla.Range.Cells
            ' Verificar si la celda contiene una tabla
            If celda.Tables.Count > 0 Then
                ' Recorrer la tabla interna dentro de la celda
                For Each tablaInterna In celda.Tables
                    tablaInterna.PreferredWidth = CentimetersToPoints(16.5)
                Next tablaInterna
            End If
        Next celda
    Next tabla
End Sub

Sub AjustaTamanoTablasAnchas()
    Dim tabla As Table
    
    ' Recorrer todas las tablas en el documento
    For Each tabla In ActiveDocument.Tables
        ' Verificar si el ancho de la tabla es mayor que 15.2 centímetros
        If tabla.Columns.Width > CentimetersToPoints(15.2) Then
            ' Ajustar el ancho de la tabla para ajustar al contenido de la ventana
            tabla.AutoFitBehavior (wdAutoFitWindow)
        End If
    Next tabla
End Sub

Function CentimetersToPoints(ByVal Centimeters As Single) As Single
    ' Convertir centímetros a puntos (1 cm = 28.35 puntos)
    CentimetersToPoints = Centimeters * 28.35
End Function
