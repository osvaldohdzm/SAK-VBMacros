Attribute VB_Name = "Módulo1"
Sub ReemplazarConURLs()
    Dim celda As Range
    Dim datos As String
    Dim resultado As String
    
    ' Rango seleccionado
    Set rangoSeleccionado = Selection
    
    ' Iterar sobre cada celda del rango
    For Each celda In rangoSeleccionado
        ' Obtener el contenido de la celda
        datos = celda.Value
        
        ' Verificar si la celda contiene una lista de Python representada como texto
        If InStr(datos, "'url':") > 0 Then
            ' Procesar solo si contiene la estructura esperada
            ' Extraer las URLs del contenido
            resultado = ConcatenarURLs(datos)
            
            ' Reemplazar el contenido de la celda con las URLs concatenadas
            celda.Value = resultado
        End If
    Next celda
End Sub

Function ConcatenarURLs(datos As String) As String
    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    Dim resultado As String
    
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.Pattern = "'url':\s*'([^']+)'"
    
    Set matches = regex.Execute(datos)
    
    ' Verificar si se encontraron coincidencias
    If matches.Count > 0 Then
        For Each match In matches
            resultado = resultado & match.SubMatches(0) & vbCrLf
        Next match
    End If
    
    ConcatenarURLs = resultado
End Function

