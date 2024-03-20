Attribute VB_Name = "Módulo1"
Sub ExtraerCadenas()

    ' Declarar variables
    Dim strDoc As String
    Dim strCadena As String
    Dim dictCampos As Object
    Set dictCampos = CreateObject("Scripting.Dictionary")
    Dim i As Long, j As Long

    ' Obtener el texto del documento
    strDoc = ActiveDocument.Content

    ' Buscar la primera ocurrencia de "«"
    i = InStr(1, strDoc, "«")

    ' Mientras haya "«"
    Do While i > 0
        ' Buscar la siguiente ocurrencia de "»"
        j = InStr(i + 1, strDoc, "»")

        ' Si se encontró "»"
        If j > 0 Then
            ' Extraer la cadena entre "«" y "»" incluyendo los caracteres
            strCadena = Mid(strDoc, i, j - i + 1)

            ' Agregar la cadena al diccionario si no está presente
            If Not dictCampos.Exists(strCadena) Then
                dictCampos.Add strCadena, strCadena
            End If

            ' Buscar la siguiente ocurrencia de "«"
            i = InStr(j + 1, strDoc, "«")
        Else
            ' No se encontró "»"
            Exit Do
        End If
    Loop

    ' Crear el archivo .txt en la misma carpeta que el documento activo
    Dim objFSO As Object
    Dim objFile As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFSO.CreateTextFile(ActiveDocument.Path & "\Campos_documento.txt")

    ' Escribir las cadenas únicas en el archivo .txt
    Dim key As Variant
    For Each key In dictCampos.Keys
        objFile.WriteLine key
    Next key

    ' Cerrar el archivo .txt
    objFile.Close

    ' Mostrar mensaje de información con la ruta completa del archivo
    MsgBox "Las cadenas se han extraído correctamente en el archivo Campos_documento.txt ubicado en: " & ActiveDocument.Path, vbInformation

End Sub

