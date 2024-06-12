Attribute VB_Name = "Módulo1"
Sub BuscarIPsDesdeArchivo()
    ' Declaración de variables
    Dim IPList As String
    Dim IPArray() As String
    Dim IP As Variant
    Dim Found As Boolean
    Dim ws As Worksheet
    Dim NewSheet As Worksheet
    Dim LastRow As Long
    Dim ResultRow As Long
    Dim FilePath As String
    Dim FileContent As String
    Dim FileLines() As String
    Dim i As Long
    
    ' Solicitar al usuario que seleccione un archivo de texto
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Seleccione el archivo de texto con la lista de IPs"
        .Filters.Clear
        .Filters.Add "Archivos de texto", "*.txt"
        .AllowMultiSelect = False
        If .Show = -1 Then
            FilePath = .SelectedItems(1)
        Else
            MsgBox "No se seleccionó ningún archivo. La operación ha sido cancelada.", vbExclamation
            Exit Sub
        End If
    End With
    
    ' Leer el contenido del archivo de texto
    On Error Resume Next
    Open FilePath For Input As #1
    FileContent = Input$(LOF(1), 1)
    Close #1
    On Error GoTo 0
    
    ' Comprobar si se ha leído correctamente el archivo
    If FileContent = "" Then
        MsgBox "No se pudo leer el archivo seleccionado. Por favor, compruebe que el archivo seleccionado contenga datos válidos.", vbCritical
        Exit Sub
    End If
    
    ' Convertir el contenido del archivo en un array de líneas
    FileLines = Split(FileContent, vbCrLf)
    
    ' Comprobar si hay líneas en el archivo
    If UBound(FileLines) = -1 Then
        MsgBox "El archivo seleccionado está vacío. Por favor, asegúrese de que el archivo contiene la lista de IPs a buscar.", vbExclamation
        Exit Sub
    End If
    
    ' Construir la lista de IPs a partir de las líneas del archivo
    For i = LBound(FileLines) To UBound(FileLines)
        IPList = IPList & FileLines(i) & " "
    Next i
    
    ' Convertir la cadena de IPs en un array utilizando el espacio como delimitador
    IPArray = Split(Trim(IPList))
    
    ' Crear una nueva hoja de cálculo para los resultados
    Set NewSheet = ThisWorkbook.Sheets.Add(After:= _
             ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    NewSheet.Name = "Resultados de búsqueda"

    ' Inicializar la fila de resultados
    ResultRow = 1

    ' Iterar sobre cada IP en el array
    For Each IP In IPArray
        ' Restablecer la bandera de encontrada
        Found = False

        ' Iterar sobre todas las hojas del libro
        For Each ws In ThisWorkbook.Sheets
            ' Buscar la IP en la hoja actual
            If WorksheetFunction.CountIf(ws.Cells, IP) > 0 Then
                ' La IP fue encontrada en la hoja actual
                Found = True
                Exit For ' Salir del bucle ya que la IP fue encontrada
            End If
        Next ws

        ' Escribir el resultado en la nueva hoja de resultados
        NewSheet.Cells(ResultRow, 1).Value = IP
        If Found Then
            NewSheet.Cells(ResultRow, 2).Value = "Encontrada"
        Else
            NewSheet.Cells(ResultRow, 2).Value = "No encontrada"
        End If

        ' Incrementar la fila de resultados
        ResultRow = ResultRow + 1
    Next IP

    ' Ajustar el ancho de las columnas en la nueva hoja de resultados
    NewSheet.Columns.AutoFit

    MsgBox "Búsqueda completada. Se han creado los resultados en una nueva hoja.", vbInformation
End Sub


