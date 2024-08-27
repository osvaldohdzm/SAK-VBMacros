Attribute VB_Name = "Módulo1"
Sub SeleccionarTablaYCompletarDatos()
    Dim ws As Worksheet
    Dim lstObjects As ListObjects
    Dim lstObject As ListObject
    Dim selectedTableName As String
    Dim selectedTable As ListObject
    Dim headerRow As Range
    Dim colTipoSolucion As Range
    Dim colTipoVulnerabilidad As Range
    Dim rowIndex As Long
    Dim correspondencia As Object
    Dim key As Variant

    ' Inicializar variables
    Set ws = ActiveSheet
    Set lstObjects = ws.ListObjects
    Set correspondencia = CreateObject("Scripting.Dictionary") ' Crear objeto Dictionary

    ' Llenar tabla de correspondencia
    correspondencia.Add "Parche de seguridad", "Ausencia de parche de seguridad"
    correspondencia.Add "Código", "Código inseguro"
    correspondencia.Add "Configuración", "Configuración insegura"
    correspondencia.Add "Actualización", "Versión desactualizada de software"
    correspondencia.Add "Versión desactualizada", "Versión desactualizada de software"

    ' Mostrar lista de selección de rango
    On Error Resume Next
    Set lstObject = Application.InputBox("Selecciona una celda dentro de la tabla", Type:=8).ListObject
    On Error GoTo 0

    ' Verificar si se seleccionó una tabla válida
    If lstObject Is Nothing Then
        MsgBox "No se seleccionó una tabla válida.", vbExclamation
        Exit Sub
    End If

    ' Establecer la tabla seleccionada
    Set selectedTable = lstObject

    ' Verificar la existencia de las columnas TipoSolucion y TipoVulnerabilidad
    On Error Resume Next
    Set colTipoSolucion = selectedTable.ListColumns("TipoSolucion").DataBodyRange
    Set colTipoVulnerabilidad = selectedTable.ListColumns("TipoVulnerabilidad").DataBodyRange
    On Error GoTo 0

    If colTipoSolucion Is Nothing Or colTipoVulnerabilidad Is Nothing Then
        MsgBox "No se encontró la columna 'TipoSolucion' o 'TipoVulnerabilidad'. La macro no puede continuar.", vbExclamation
        Exit Sub
    End If

    ' Rellenar los valores vacíos en la columna TipoVulnerabilidad
    For rowIndex = 1 To colTipoSolucion.Rows.Count
        If colTipoVulnerabilidad.Cells(rowIndex, 1).Value = "" Then
            key = colTipoSolucion.Cells(rowIndex, 1).Value
            If correspondencia.Exists(key) Then
                colTipoVulnerabilidad.Cells(rowIndex, 1).Value = correspondencia(key)
            End If
        End If
    Next rowIndex

    MsgBox "Proceso completado.", vbInformation
End Sub

