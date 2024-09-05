Attribute VB_Name = "Module1"
Sub LimpiarEtiquetasHTML()
    Dim selectedRange As Range
    Dim cell As Range
    Dim htmlPattern As String
    
    ' Definir el patrón HTML que se desea eliminar
    htmlPattern = "<(\/?(p|a|li|ul|b|strong|i|u|br)[^>]*?)>|<\/p><p>"
    
    ' Obtener el rango de celdas seleccionadas por el usuario
    On Error Resume Next
    Set selectedRange = Application.InputBox("Seleccione el rango de celdas", Type:=8)
    On Error GoTo 0
    
    ' Salir si el usuario cancela la selección
    If selectedRange Is Nothing Then Exit Sub
    
    ' Iterar sobre cada celda en el rango seleccionado
    For Each cell In selectedRange
        ' Verificar si la celda contiene texto
        If Not IsEmpty(cell.Value) And TypeName(cell.Value) = "String" Then
            ' Eliminar las etiquetas HTML utilizando expresiones regulares
            cell.Value = RegExpReplace(cell.Value, htmlPattern, vbCrLf) ' Reemplazar con salto de línea
        End If
    Next cell
    
    MsgBox "Etiquetas HTML eliminadas correctamente y reemplazadas según lo solicitado.", vbInformation
End Sub

Function RegExpReplace(ByVal text As String, ByVal replacePattern As String, ByVal replaceWith As String) As String
    ' Función para reemplazar utilizando expresiones regulares
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    With regex
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .Pattern = replacePattern
    End With
    
    RegExpReplace = regex.Replace(text, replaceWith)
End Function

