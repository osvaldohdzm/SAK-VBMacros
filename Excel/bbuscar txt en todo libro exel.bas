Attribute VB_Name = "MÃ³dulo1"
Sub SearchStringsInAllSheets()
    Dim filePath As String
    Dim fileContent As String
    Dim searchTerms As Variant
    Dim term As Variant
    Dim ws As Worksheet
    Dim cell As Range
    Dim found As Boolean
    Dim results As Worksheet
    Dim resultRow As Long

    ' Solicitar archivo de texto
    filePath = Application.GetOpenFilename("Text Files (*.txt), *.txt", , "Selecciona un archivo de texto")
    If filePath = "False" Then Exit Sub ' Si el usuario cancela, salir

    ' Leer el contenido del archivo
    Dim fileNum As Integer
    fileNum = FreeFile
    On Error GoTo FileReadError
    Open filePath For Input As fileNum
    fileContent = Input$(LOF(fileNum), fileNum)
    Close fileNum
    On Error GoTo 0

    ' Separar el contenido del archivo en líneas
    searchTerms = Split(fileContent, vbCrLf)

    ' Verificar si ya existe una hoja de resultados y eliminarla
    On Error Resume Next
    Set results = ThisWorkbook.Worksheets("Resultados")
    If Not results Is Nothing Then results.Delete
    On Error GoTo 0

    ' Crear una nueva hoja para los resultados
    Set results = Worksheets.Add
    results.Name = "Resultados"
    results.Cells(1, 1).Value = "Cadena de Búsqueda"
    results.Cells(1, 2).Value = "Resultado"
    resultRow = 2

    ' Realizar la búsqueda para cada término
    For Each term In searchTerms
        found = False
        ' Buscar en todas las celdas de cada hoja del libro
        For Each ws In ThisWorkbook.Worksheets
            If ws.Name <> "Resultados" Then ' Omitir la hoja de resultados en la búsqueda
                For Each cell In ws.UsedRange
                    If InStr(1, cell.Value, term, vbTextCompare) > 0 Then
                        found = True
                        Exit For
                    End If
                Next cell
            End If
            If found Then Exit For
        Next ws

        ' Registrar el resultado en la hoja "Resultados"
        results.Cells(resultRow, 1).Value = term
        If found Then
            results.Cells(resultRow, 2).Value = "FOUND"
        Else
            results.Cells(resultRow, 2).Value = "NOT FOUND"
        End If
        resultRow = resultRow + 1
    Next term

    MsgBox "Búsqueda completada. Resultados en la hoja 'Resultados'.", vbInformation
    Exit Sub

FileReadError:
    MsgBox "Error al leer el archivo. Verifica que el archivo esté disponible y sea válido.", vbCritical
    If fileNum <> 0 Then Close fileNum
End Sub

