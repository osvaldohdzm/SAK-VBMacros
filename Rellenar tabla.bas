Attribute VB_Name = "Módulo1"
Sub RellenarValores()
    Dim ws As Worksheet
    Dim selectedTable As ListObject
    Dim colTipoSolucion As ListColumn
    Dim colTipoVulnerabilidad As ListColumn
    Dim rowIndex As Long
    Dim key As Variant
    Dim headerRange As Range
    Dim colTipoSolucionHeader As Range
    Dim colTipoVulnerabilidadHeader As Range
    Dim correspondencia As Object

    ' Inicializar variables
    Set ws = ActiveSheet
    Set correspondencia = CreateObject("Scripting.Dictionary") ' Crear objeto Dictionary
    
    ' Llenar tabla de correspondencia
    correspondencia.Add "Parche de seguridad", "Ausencia de parche de seguridad"
    correspondencia.Add "Código", "Código inseguro"
    correspondencia.Add "Configuración", "Configuración insegura"
    correspondencia.Add "Actualización", "Versión desactualizada de software"
    correspondencia.Add "Versión desactualizada", "Versión desactualizada de software"
    
    ' Solicitar al usuario seleccionar la celda del encabezado de la columna TipoSolucion
    On Error Resume Next
    Set headerRange = Application.InputBox("Selecciona la celda del encabezado de 'TipoSolucion'", Type:=8)
    On Error GoTo 0

    If headerRange Is Nothing Then
        MsgBox "No se seleccionó ningún encabezado. La macro no puede continuar.", vbExclamation
        Exit Sub
    End If
    
    ' Verificar que la celda seleccionada pertenece a una tabla
    If headerRange.ListObject Is Nothing Then
        MsgBox "La celda seleccionada no pertenece a una tabla. La macro no puede continuar.", vbExclamation
        Exit Sub
    End If
    
    Set selectedTable = headerRange.ListObject

    ' Obtener la columna TipoSolucion
    Set colTipoSolucionHeader = headerRange
    Set colTipoSolucion = selectedTable.ListColumns(colTipoSolucionHeader.Value)

    ' Solicitar al usuario seleccionar la celda del encabezado de la columna TipoVulnerabilidad
    On Error Resume Next
    Set headerRange = Application.InputBox("Selecciona la celda del encabezado de 'TipoVulnerabilidad'", Type:=8)
    On Error GoTo 0

    If headerRange Is Nothing Then
        MsgBox "No se seleccionó ningún encabezado. La macro no puede continuar.", vbExclamation
        Exit Sub
    End If

    ' Verificar que la celda seleccionada pertenece a una tabla
    If headerRange.ListObject Is Nothing Then
        MsgBox "La celda seleccionada no pertenece a una tabla. La macro no puede continuar.", vbExclamation
        Exit Sub
    End If

    ' Verificar que la tabla seleccionada es la misma que la primera
    If headerRange.ListObject.Name <> selectedTable.Name Then
        MsgBox "La celda seleccionada no pertenece a la misma tabla. La macro no puede continuar.", vbExclamation
        Exit Sub
    End If

    ' Obtener la columna TipoVulnerabilidad
    Set colTipoVulnerabilidadHeader = headerRange
    Set colTipoVulnerabilidad = selectedTable.ListColumns(colTipoVulnerabilidadHeader.Value)

    ' Rellenar los valores vacíos en la columna TipoVulnerabilidad
    For rowIndex = 1 To selectedTable.ListRows.Count
        If colTipoVulnerabilidad.DataBodyRange.Cells(rowIndex, 1).Value = "" Then
            key = colTipoSolucion.DataBodyRange.Cells(rowIndex, 1).Value
            If correspondencia.Exists(key) Then
                colTipoVulnerabilidad.DataBodyRange.Cells(rowIndex, 1).Value = correspondencia(key)
            End If
        End If
    Next rowIndex

    MsgBox "Proceso completado.", vbInformation
End Sub

