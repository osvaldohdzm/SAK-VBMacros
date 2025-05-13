Attribute VB_Name = "Módulo1"
Sub FiltrarPorValores()

    Dim ListaValores As String
    Dim Encabezado As Range
    Dim Filtro() As String
    Dim ColIndex As Integer
    Dim UltimaFila As Long
    
    ' Solicitar lista de valores al usuario
    ListaValores = InputBox("Ingrese la lista de valores separados por comas:", "Lista de Valores")
    
    ' Salir si el usuario cancela
    If ListaValores = "" Then Exit Sub
    
    ' Convertir la lista de valores en un array
    Filtro = Split(ListaValores, ",")
    
    ' Solicitar al usuario seleccionar el encabezado
    On Error Resume Next
    Set Encabezado = Application.InputBox("Seleccione el encabezado de la columna donde aplicar el filtro:", Type:=8)
    On Error GoTo 0
    
    ' Salir si el usuario cancela o no selecciona un encabezado válido
    If Encabezado Is Nothing Then Exit Sub
    
    ' Determinar el índice de la columna
    ColIndex = Encabezado.Column
    
    ' Determinar la última fila de la columna
    UltimaFila = Encabezado.Offset(ActiveSheet.Rows.Count - Encabezado.Row, 0).End(xlUp).Row
    
    ' Desactivar el modo de autofiltro para eliminar cualquier filtro existente
    If ActiveSheet.AutoFilterMode Then
        ActiveSheet.AutoFilterMode = False
    End If
    
    ' Aplicar el filtro en la columna correspondiente
    Encabezado.Parent.Range(Encabezado.Offset(1, 0), Encabezado.Offset(UltimaFila - Encabezado.Row, 0)).AutoFilter Field:=ColIndex, Criteria1:=Filtro, Operator:=xlFilterValues
    
    ' Informar al usuario sobre la finalización del proceso
    MsgBox "El filtro se ha aplicado correctamente.", vbInformation

End Sub

