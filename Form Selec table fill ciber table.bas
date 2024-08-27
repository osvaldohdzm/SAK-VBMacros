Attribute VB_Name = "Módulo1"
Option Explicit

Dim selectedTableName As String
Dim correspondencia As Object

Sub MostrarFormularioSeleccionarTabla()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim frm As Object
    Dim lstTablas As MSForms.ListBox
    Dim btnAceptar As MSForms.CommandButton
    Dim btnCancelar As MSForms.CommandButton
    
    Set frm = VBA.UserForms.Add("UserForm1")
    
    ' Configurar UserForm
    With frm
        .Caption = "Seleccionar Tabla"
        .Width = 300
        .Height = 200
    End With
    
    ' Agregar ListBox
    Set lstTablas = frm.Controls.Add("Forms.ListBox.1", "lstTablas", True)
    With lstTablas
        .Top = 10
        .Left = 10
        .Width = 200
        .Height = 120
    End With
    
    ' Llenar el ListBox con los nombres de las tablas de la hoja activa
    Set ws = ActiveSheet
    For Each tbl In ws.ListObjects
        lstTablas.AddItem tbl.Name
    Next tbl
    
    ' Manejar eventos de los botones
    frm.Show vbModal
    
    ' Manejo del evento de clic en el botón Aceptar
    If frm.Controls("lstTablas").ListIndex <> -1 Then
        selectedTableName = frm.Controls("lstTablas").Value
        Unload frm
        Call RellenarValores
    Else
        MsgBox "No se seleccionó una tabla válida.", vbExclamation
        Unload frm
    End If
End Sub

Sub RellenarValores()
    Dim ws As Worksheet
    Dim selectedTable As ListObject
    Dim colTipoSolucion As ListColumn
    Dim colTipoVulnerabilidad As ListColumn
    Dim rowIndex As Long
    Dim key As Variant
    
    ' Inicializar variables
    Set ws = ActiveSheet
    Set correspondencia = CreateObject("Scripting.Dictionary") ' Crear objeto Dictionary
    
    ' Llenar tabla de correspondencia
    correspondencia.Add "Parche de seguridad", "Ausencia de parche de seguridad"
    correspondencia.Add "Código", "Código inseguro"
    correspondencia.Add "Configuración", "Configuración insegura"
    correspondencia.Add "Actualización", "Versión desactualizada de software"
    correspondencia.Add "Versión desactualizada", "Versión desactualizada de software"
    
    ' Buscar la tabla seleccionada por el nombre
    On Error Resume Next
    Set selectedTable = ws.ListObjects(selectedTableName)
    On Error GoTo 0
    
    If selectedTable Is Nothing Then
        MsgBox "No se encontró la tabla seleccionada.", vbExclamation
        Exit Sub
    End If

    ' Verificar la existencia de las columnas TipoSolucion y TipoVulnerabilidad
    On Error Resume Next
    Set colTipoSolucion = selectedTable.ListColumns("TipoSolucion")
    Set colTipoVulnerabilidad = selectedTable.ListColumns("TipoVulnerabilidad")
    On Error GoTo 0
    
    If colTipoSolucion Is Nothing Or colTipoVulnerabilidad Is Nothing Then
        MsgBox "No se encontró la columna 'TipoSolucion' o 'TipoVulnerabilidad'. La macro no puede continuar.", vbExclamation
        Exit Sub
    End If

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

