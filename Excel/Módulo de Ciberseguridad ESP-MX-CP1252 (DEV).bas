Attribute VB_Name = "ExcelModuloCiberseguridad"
Sub EliminarLineasVaciasEnCelda(WordDoc As Object)
    Dim rng As Object ' Usando tipo Object en lugar de Word.Range
    Dim textoOriginal As String

    ' Asegurarse de que el documento no sea nulo
    If WordDoc Is Nothing Then
        MsgBox "El documento es nulo.", vbExclamation
        Exit Sub
    End If

    ' Obtener la celda específica: WordDoc.Tables(1).Cell(8, 1)
    Dim cell As Object
    Set cell = WordDoc.Tables(1).cell(8, 1)

    ' La comprobación TypeOf ya no es estrictamente necesaria si declaras cell As Object,
    ' pero la dejamos por si acaso se pasa un Nothing.
    If cell Is Nothing Then
        MsgBox "Se recibió un objeto de celda nulo.", vbExclamation
        Exit Sub
    End If

    ' Intentar asignar el rango de la celda
    On Error Resume Next
    Set rng = cell.Range
    On Error GoTo 0

    ' Verificar si el rango fue correctamente asignado
    If rng Is Nothing Then
        MsgBox "No se pudo obtener el rango de la celda.", vbExclamation
        Exit Sub
    End If

    ' Asegurarse de que el rango contiene texto antes de realizar cambios
    If Len(rng.text) > 0 Then
        textoOriginal = rng.text
        Debug.Print "Texto original (longitud " & Len(textoOriginal) & "): [" & textoOriginal & "]" ' Depuración

        ' Verificar si el último carácter del RANGO es el marcador de fin de celda (Chr(7))
        On Error Resume Next
        If rng.Characters.Last.text = Chr(7) Then
            If Err.Number = 0 Then ' Solo proceder si no hubo error al acceder a Last.Text
                ' Si es Chr(7), ajustamos el RANGO para excluirlo.
                rng.End = rng.End - 1
                Debug.Print "Ajustado el rango para excluir Chr(7). Nuevo final: " & rng.End
                ' Volver a verificar si el rango aún tiene longitud después del ajuste
                If rng.Start >= rng.End Then ' Cambiado a >= por seguridad
                    Debug.Print "El rango quedó vacío después de quitar Chr(7). Saliendo."
                    On Error GoTo 0 ' Restablecer manejo de errores
                    Exit Sub ' No hay nada más que hacer si el rango está vacío
                End If
            Else
                Debug.Print "Error accediendo a rng.Characters.Last: " & Err.Description
                Err.Clear
            End If
        End If
        On Error GoTo 0 ' Restablecer manejo de errores normal

        ' Reemplazar saltos de párrafo múltiples (^13) con uno solo usando comodines
        With rng.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .text = "(^13){2,}"      ' Buscar DOS O MÁS saltos de párrafo (^13) consecutivos
            .Replacement.text = "^13" ' Reemplazar con UN SOLO salto de párrafo (^13) cuando se usan comodines
            .Forward = True
            .Wrap = wdFindStop      ' No continuar la búsqueda al inicio si llega al final
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = True  ' ¡IMPORTANTE! Habilitar comodines
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .Execute Replace:=wdReplaceAll ' Ejecutar el reemplazo en todo el rango (ajustado)
        End With

        ' --- Opcional: Eliminar saltos de párrafo/línea al principio del rango ---
        Do While rng.Characters.Count > 0 And (Left(rng.text, 1) = vbCr Or Left(rng.text, 1) = vbLf)
            rng.Start = rng.Start + 1 ' Mover el inicio del rango un carácter hacia adelante
        Loop
        Debug.Print "Texto después de quitar saltos iniciales: [" & rng.text & "]"

        ' --- Opcional: Eliminar un único salto de párrafo/línea al final del rango ---
        If rng.Characters.Count > 0 Then ' Asegurarse que el rango no está vacío
            Dim ultimoChar As String
            ultimoChar = rng.Characters.Last.text
            If ultimoChar = Chr(13) Or ultimoChar = Chr(10) Then ' Chr(13)=vbCr, Chr(10)=vbLf
                rng.End = rng.End - 1 ' Ajustar el final del rango para eliminarlo
                Debug.Print "Eliminado salto de párrafo/línea final. Nuevo final: " & rng.End
            End If
        End If

        Debug.Print "Texto final en rango (longitud " & Len(rng.text) & "): [" & rng.text & "]"
    Else
        Debug.Print "La celda original estaba vacía o el rango quedó vacío."
    End If

    ' Limpiar la referencia al rango después de finalizar
    Set rng = Nothing
End Sub


Sub CYB008_LimpiarTextoYAgregarGuion()
    Dim celda As Range
    Dim Texto As String
    Dim textoLimpio As String
    Dim lineas As Variant
    Dim i As Integer
    Dim textoConGuiones As String
    Dim incluirGuion As Boolean
    
    ' Recorremos cada celda seleccionada
    For Each celda In Selection
        ' Solo procesamos celdas con texto
        If Not IsEmpty(celda.value) Then
            Texto = celda.value
            
            ' Mantener las líneas vacías pero eliminar saltos de línea innecesarios dentro del texto
            lineas = Split(Texto, vbLf)
            textoLimpio = ""
            
            ' Eliminar las líneas vacías (CHAR(10)) pero mantener los saltos de línea necesarios
            For i = LBound(lineas) To UBound(lineas)
                If Len(Trim(lineas(i))) > 0 Then
                    textoLimpio = textoLimpio & lineas(i) & vbLf
                End If
            Next i
            
            ' Eliminar el salto de línea final extra
            If Len(textoLimpio) > 0 Then
                textoLimpio = Left(textoLimpio, Len(textoLimpio) - 1)
            End If
            
            ' Inicializamos la variable para el texto con guiones
            textoConGuiones = ""
            lineas = Split(textoLimpio, vbLf)
            incluirGuion = False
            
            ' Recorrer las líneas y agregar guiones a partir del primer ":"
            For i = LBound(lineas) To UBound(lineas)
                If InStr(1, lineas(i), ":", vbTextCompare) > 0 And Not incluirGuion Then
                    ' Agregamos la línea con los ":" pero sin gui?n
                    textoConGuiones = textoConGuiones & lineas(i) & vbLf
                    incluirGuion = True ' Habilitamos la adici?n de guiones despu?s de encontrar el ":"
                ElseIf incluirGuion Then
                    ' Despu?s del primer ":", agregamos un guion
                    If Len(Trim(lineas(i))) > 0 Then
                        textoConGuiones = textoConGuiones & " - " & lineas(i) & vbLf
                    Else
                        ' Si la línea est? vacía, solo agregamos el salto de línea
                        textoConGuiones = textoConGuiones & vbLf
                    End If
                Else
                    ' Si aún no hemos encontrado el ":", no agregamos guiones
                    textoConGuiones = textoConGuiones & lineas(i) & vbLf
                End If
            Next i
            
            ' Eliminar el salto de línea final extra
            If Len(textoConGuiones) > 0 Then
                textoConGuiones = Left(textoConGuiones, Len(textoConGuiones) - 1)
            End If
            
            ' Asignar el texto limpio con guiones de vuelta a la celda
            celda.value = textoConGuiones
        End If
    Next celda
End Sub







Sub CYB020_ExportarHojaConFormatoINAI()
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim tempFileName As String
    Dim carpetaSalida As String
    Dim tbl As ListObject
    Dim colSeveridad As ListColumn
    Dim selectedRange As Range
    
    ' Mostrar cuadro de di?logo para seleccionar la carpeta de destino
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Selecciona la carpeta de salida"
        If .Show = -1 Then
            carpetaSalida = .SelectedItems(1)
        Else
            MsgBox "No se seleccion? ninguna carpeta.", vbExclamation
            Exit Sub
        End If
    End With
    
    ' Exportar la hoja activa a un archivo Excel
    Set ws = ActiveSheet
    If Not ws Is Nothing Then
        tempFileName = carpetaSalida & "\" & "SSIFO37-02_Matriz de seguimiento vulnerabilidades de aplicaciones.xlsx"
        ws.Copy
        Set wb = ActiveWorkbook
        ' Aplicar formato a la hoja exportada
        With wb.Sheets(1)
            ' Ajustar altura de las filas
            .Cells.Select
            Selection.RowHeight = 15
            
            ' Centrar la columna A (ajustar según las necesidades)
            Columns("A:A").Select
            With Selection
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
            End With
            
            ' Centrar la columna C y "Severidad"
            Columns("C:C").Select
            With Selection
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
            
            ' Aplicar el estilo de tabla a la primera tabla
            .ListObjects(1).TableStyle = "TableStyleMedium1"
            
            ' Buscar la columna "Severidad" en la primera tabla
            Set tbl = .ListObjects(1)
            On Error Resume Next
            Set colSeveridad = tbl.ListColumns("Severidad")
            On Error GoTo 0
            
            ' Verificar si se encontr? la columna "Severidad"
            If Not colSeveridad Is Nothing Then
                ' Aplicar formato condicional a la columna "Severidad"
                Set selectedRange = colSeveridad.DataBodyRange
                With selectedRange
                    ' CRÍTICA
                    .FormatConditions.Add Type:=xlTextString, String:="CRÍTICA", TextOperator:=xlContains
                    .FormatConditions(.FormatConditions.Count).SetFirstPriority
                    With .FormatConditions(1).Font
                        .Color = RGB(255, 255, 255)        ' Blanco
                    End With
                    With .FormatConditions(1).Interior
                        .Color = RGB(112, 48, 160)        ' #7030A0
                    End With
                    
                    ' ALTA
                    .FormatConditions.Add Type:=xlTextString, String:="ALTA", TextOperator:=xlContains
                    .FormatConditions(.FormatConditions.Count).SetFirstPriority
                    With .FormatConditions(1).Font
                        .Color = RGB(255, 255, 255)        ' Blanco
                    End With
                    With .FormatConditions(1).Interior
                        .Color = RGB(255, 0, 0)        ' #FF0000
                    End With
                    
                    ' MEDIA
                    .FormatConditions.Add Type:=xlTextString, String:="MEDIA", TextOperator:=xlContains
                    .FormatConditions(.FormatConditions.Count).SetFirstPriority
                    With .FormatConditions(1).Font
                        .Color = RGB(0, 0, 0)        ' Negro
                    End With
                    With .FormatConditions(1).Interior
                        .Color = RGB(255, 255, 0)        ' #FFFF00
                    End With
                    
                    ' BAJA
                    .FormatConditions.Add Type:=xlTextString, String:="BAJA", TextOperator:=xlContains
                    .FormatConditions(.FormatConditions.Count).SetFirstPriority
                    With .FormatConditions(1).Font
                        .Color = RGB(255, 255, 255)        ' Blanco
                    End With
                    With .FormatConditions(1).Interior
                        .Color = RGB(0, 176, 80)        ' #00B050
                    End With
                    
                    ' INFORMATIVA
                    .FormatConditions.Add Type:=xlTextString, String:="INFORMATIVA", TextOperator:=xlContains
                    .FormatConditions(.FormatConditions.Count).SetFirstPriority
                    With .FormatConditions(1).Font
                        .Color = RGB(0, 0, 0)        ' Negro
                    End With
                    With .FormatConditions(1).Interior
                        .Color = RGB(231, 230, 230)        ' #E7E6E6
                    End With
                End With
            Else
                MsgBox "No se encontr? la columna        'Severidad'.", vbExclamation
            End If
        End With
        
        wb.Sheets(1).Name = "Vulnerabilidades"
        wb.Sheets(1).Range("A1").Select
        ' Guardar el archivo en la carpeta seleccionada
        wb.SaveAs tempFileName, xlOpenXMLWorkbook
        wb.Close False
    End If
    
End Sub

Function FunExportarHojaActivaAExcelINAI(carpetaSalida As String, folderName As String, ws As Worksheet, fileName As String) As Boolean
    Dim wb          As Workbook
    Dim tempFileName As String
    Dim tbl         As ListObject
    Dim colSeveridad As ListColumn
    Dim selectedRange As Range
    
    On Error GoTo ErrorHandler
    
    ' Exportar la hoja activa a un archivo Excel
    ws.Activate
    If Not ws Is Nothing Then
        tempFileName = carpetaSalida & "\" & fileName & ".xlsx"
        ws.Copy
        Set wb = ActiveWorkbook
        
        ' Aplicar formato a la hoja exportada
        With wb.Sheets(1)
            ' Ajustar altura de las filas
            .Cells.RowHeight = 15
            
            ' Centrar la columna A
            .Columns("A:A").HorizontalAlignment = xlCenter
            .Columns("A:A").VerticalAlignment = xlBottom
            
            ' Centrar la columna C y "Severidad"
            .Columns("C:C").HorizontalAlignment = xlCenter
            .Columns("C:C").VerticalAlignment = xlCenter
            
            ' Aplicar el estilo de tabla a la primera tabla
            .ListObjects(1).TableStyle = "TableStyleMedium1"
            
            ' Buscar la columna "Severidad" en la primera tabla
            Set tbl = .ListObjects(1)
            On Error Resume Next
            Set colSeveridad = tbl.ListColumns("Severidad")
            On Error GoTo 0
            
            ' Verificar si se encontr? la columna "Severidad"
            If Not colSeveridad Is Nothing Then
                ' Aplicar formato condicional a la columna "Severidad"
                Set selectedRange = colSeveridad.DataBodyRange
                With selectedRange
                    ' CRÍTICA
                    .FormatConditions.Add Type:=xlTextString, String:="CRÍTICA", TextOperator:=xlContains
                    .FormatConditions(.FormatConditions.Count).SetFirstPriority
                    With .FormatConditions(1).Font
                        .Color = RGB(255, 255, 255)        ' Blanco
                    End With
                    With .FormatConditions(1).Interior
                        .Color = RGB(112, 48, 160)        ' #7030A0
                    End With
                    
                    ' ALTA
                    .FormatConditions.Add Type:=xlTextString, String:="ALTA", TextOperator:=xlContains
                    .FormatConditions(.FormatConditions.Count).SetFirstPriority
                    With .FormatConditions(1).Font
                        .Color = RGB(255, 255, 255)        ' Blanco
                    End With
                    With .FormatConditions(1).Interior
                        .Color = RGB(255, 0, 0)        ' #FF0000
                    End With
                    
                    ' MEDIA
                    .FormatConditions.Add Type:=xlTextString, String:="MEDIA", TextOperator:=xlContains
                    .FormatConditions(.FormatConditions.Count).SetFirstPriority
                    With .FormatConditions(1).Font
                        .Color = RGB(0, 0, 0)        ' Negro
                    End With
                    With .FormatConditions(1).Interior
                        .Color = RGB(255, 255, 0)        ' #FFFF00
                    End With
                    
                    ' BAJA
                    .FormatConditions.Add Type:=xlTextString, String:="BAJA", TextOperator:=xlContains
                    .FormatConditions(.FormatConditions.Count).SetFirstPriority
                    With .FormatConditions(1).Font
                        .Color = RGB(255, 255, 255)        ' Blanco
                    End With
                    With .FormatConditions(1).Interior
                        .Color = RGB(0, 176, 80)        ' #00B050
                    End With
                    
                    ' INFORMATIVA
                    .FormatConditions.Add Type:=xlTextString, String:="INFORMATIVA", TextOperator:=xlContains
                    .FormatConditions(.FormatConditions.Count).SetFirstPriority
                    With .FormatConditions(1).Font
                        .Color = RGB(0, 0, 0)        ' Negro
                    End With
                    With .FormatConditions(1).Interior
                        .Color = RGB(231, 230, 230)        ' #E7E6E6
                    End With
                End With
            Else
                MsgBox "No se encontr? la columna        'Severidad'.", vbExclamation
            End If
        End With
        
        wb.Sheets(1).Name = "Vulnerabilidades"
        wb.Sheets(1).Range("A1").Select
        ' Guardar el archivo en la carpeta seleccionada
        wb.SaveAs tempFileName, xlOpenXMLWorkbook
        wb.Close False
        
        FunExportarHojaActivaAExcelINAI = True
        MsgBox "La hoja ha sido exportada con ?xito a " & tempFileName, vbInformation
    Else
        FunExportarHojaActivaAExcelINAI = False
        MsgBox "No hay ninguna hoja activa para exportar.", vbExclamation
    End If
    
    Exit Function
    
ErrorHandler:
    FunExportarHojaActivaAExcelINAI = False
    MsgBox "Ocurri? un error: " & Err.Description, vbCritical
End Function

Sub CYB032_ReemplazarCadenasSeveridades()
    
    Dim c           As Range
    Dim valorActual As String
    
    For Each c In Selection
        valorActual = Trim(UCase(c.value))        ' Convertimos a mayúsculas y eliminamos espacios adicionales
        
        Select Case valorActual
            Case "0", "NONE", "INFORMATIVA", "INFO"
                c.value = "INFORMATIVA"
            Case "1", "BAJA", "BAJO", "LOW"
                c.value = "BAJA"
            Case "2", "BAJA", "BAJO", "LOW"
                c.value = "BAJA"
            Case "3", "BAJA", "BAJO", "LOW"
                c.value = "BAJA"
            Case "4", "BAJA", "BAJO", "LOW"
                c.value = "BAJA"
            Case "5", "MEDIA", "MEDIO", "MEDIUM"
                c.value = "MEDIA"
            Case "6", "MEDIA", "MEDIO", "MEDIUM"
                c.value = "MEDIA"
            Case "7", "ALTO", "HIGH"
                c.value = "ALTA"
            Case "8", "ALTA", "HIGH"
                c.value = "ALTA"
            Case "9", "CRÍTICA", "CRITICAL", "CRÍTICO"
                c.value = "CRÍTICA"
            Case "10", "CRÍTICA", "CRITICAL", "CRÍTICO"
                c.value = "CRÍTICA"
                ' Mantener el contenido actual si no coincide con las palabras a reemplazar
            Case Else
                ' No hacer nada
        End Select
    Next c
End Sub

Sub CYB024_LimpiarCeldasYMostrarContenidoComoArray()
    Dim rng         As Range
    Dim cell        As Range
    Dim content     As String
    Dim contentArray() As String
    Dim i           As Integer
    Dim temp        As String
    Dim uniqueUrls  As Object
    Dim uniqueArray() As String
    Dim n           As Integer
    
    ' Selecciona el rango deseado
    Set rng = Selection
    
    ' Recorre cada celda en el rango
    For Each cell In rng
        ' Obtiene el contenido de la celda
        content = cell.value
        
        ' Comprueba si el contenido es vacío
        If content <> "" Then
            ' Convierte el contenido en un array separado por el car?cter de nueva línea
            contentArray = Split(content, Chr(10))
            
            ' Inicializa el diccionario para almacenar las URL únicas
            Set uniqueUrls = CreateObject("Scripting.Dictionary")
            
            ' Agrega las URL únicas al diccionario
            For i = LBound(contentArray) To UBound(contentArray)
                If contentArray(i) <> "" Then
                    ' Elimina espacios en blanco, Chr(10) y Chr(13) del elemento
                    contentArray(i) = Trim(Replace(contentArray(i), Chr(10), ""))
                    contentArray(i) = Trim(Replace(contentArray(i), Chr(13), ""))
                    contentArray(i) = Replace(contentArray(i), " ", "")
                    If InStr(1, contentArray(i), "wikipedia", vbTextCompare) = 0 Then
                        If Not uniqueUrls.Exists(contentArray(i)) Then
                            uniqueUrls.Add contentArray(i), Nothing
                        End If
                    End If
                End If
            Next i
            
            ' Convertir la colecci?n de claves en un array
            n = uniqueUrls.Count - 1
            ReDim uniqueArray(n)
            i = 0
            For Each key In uniqueUrls.Keys
                uniqueArray(i) = key
                i = i + 1
            Next
            
            ' Ordena los elementos del array
            For i = LBound(uniqueArray) To UBound(uniqueArray) - 1
                For j = i + 1 To UBound(uniqueArray)
                    If uniqueArray(i) > uniqueArray(j) Then
                        temp = uniqueArray(i)
                        uniqueArray(i) = uniqueArray(j)
                        uniqueArray(j) = temp
                    End If
                Next j
            Next i
            
            ' Convierte el array nuevamente en una cadena concatenada por el car?cter de nueva línea
            content = Join(uniqueArray, Chr(10))
            
            ' Asigna el contenido filtrado a la celda
            cell.value = content
        End If
    Next cell
End Sub

Sub CYB029_ReemplazarConURLs()
    Dim cell        As Range
    Dim parts       As Variant
    Dim url         As String
    Dim i           As Integer
    
    ' Recorre cada celda en el rango seleccionado
    For Each cell In Selection
        If cell.value <> "" Then
            ' Separa la cadena por comas
            parts = Split(cell.value, ",")
            
            ' Inicializa una cadena vacía para las URLs
            url = ""
            
            ' Recorre cada parte de la cadena
            For i = LBound(parts) To UBound(parts)
                ' Separa cada parte por el símbolo |
                If InStr(parts(i), "|") > 0 Then
                    url = url & Mid(parts(i), InStr(parts(i), "|") + 1) & vbLf
                End If
            Next i
            
            ' Elimina el último salto de línea
            If Len(url) > 0 Then
                url = Left(url, Len(url) - 1)
            End If
            
            ' Reemplaza las comillas dobles sobrantes
            url = Replace(url, """", "")
            
            ' Reemplaza el valor de la celda con las URLs y saltos de línea
            cell.value = url
        End If
    Next cell
End Sub

Sub CYB037_AplicarFormatoCondicional()
    Dim selectedRange As Range
    
    ' Verificar si hay celdas seleccionadas
    On Error Resume Next
    Set selectedRange = Selection.SpecialCells(xlCellTypeConstants)
    On Error GoTo 0
    
    ' Salir si no hay celdas seleccionadas
    If selectedRange Is Nothing Then
        MsgBox "No hay celdas seleccionadas."
        Exit Sub
    End If
    
    ' Aplicar formato condicional según el contenido de las celdas seleccionadas
    With selectedRange
        .FormatConditions.Add Type:=xlTextString, String:="CRÍTICA", TextOperator:=xlContains
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Font
            .Color = RGB(255, 255, 255)        ' Blanco
        End With
        With .FormatConditions(1).Interior
            .Color = RGB(112, 48, 160)        ' #7030A0
        End With
        
        .FormatConditions.Add Type:=xlTextString, String:="ALTA", TextOperator:=xlContains
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Font
            .Color = RGB(255, 255, 255)        ' Blanco
        End With
        With .FormatConditions(1).Interior
            .Color = RGB(255, 0, 0)        ' #FF0000
        End With
        
        .FormatConditions.Add Type:=xlTextString, String:="MEDIA", TextOperator:=xlContains
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Font
            .Color = RGB(0, 0, 0)        ' Negro
        End With
        With .FormatConditions(1).Interior
            .Color = RGB(255, 255, 0)        ' #FFFF00
        End With
        
        .FormatConditions.Add Type:=xlTextString, String:="BAJA", TextOperator:=xlContains
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Font
            .Color = RGB(255, 255, 255)        ' Blanco
        End With
        With .FormatConditions(1).Interior
            .Color = RGB(0, 176, 80)        ' #00B050
        End With
        
        .FormatConditions.Add Type:=xlTextString, String:="INFORMATIVA", TextOperator:=xlContains
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Font
            .Color = RGB(0, 0, 0)        ' Negro
        End With
        With .FormatConditions(1).Interior
            .Color = RGB(231, 230, 230)        ' #E7E6E6
        End With
    End With
End Sub

Sub CYB033_ConvertirATextoEnOracion()
    Dim celda       As Range
    Dim Texto       As String
    Dim primeraLetra As String
    Dim restoTexto  As String
    
    ' Recorre cada celda en el rango seleccionado
    For Each celda In Selection
        If Not IsEmpty(celda.value) Then
            Texto = celda.value
            ' Convierte todo el texto a minúsculas
            Texto = LCase(Texto)
            ' Extrae la primera letra
            primeraLetra = UCase(Left(Texto, 1))
            ' Extrae el resto del texto
            restoTexto = Mid(Texto, 2)
            ' Combina la primera letra en mayúsculas con el resto del texto en minúsculas
            celda.value = primeraLetra & restoTexto
        End If
    Next celda
End Sub

Sub CYB027_QuitarEspacios()
    Dim rng         As Range
    Dim c           As Range
    
    Set rng = Selection        'asume que el rango seleccionado es el que quieres modificar
    
    For Each c In rng        'recorre cada celda del rango
        c.value = Application.Trim(c.value)        'quita los espacios de la celda
    Next c
End Sub


Sub CYB009_LimpiarSalida()
    Dim rng As Range
    Dim celda As Range
    Dim lineas As Variant
    Dim i As Integer
    Dim htmlPattern As String
    Dim liPattern As String
    Dim cleanHtmlPattern As String
    Dim NuevoTexto As String
    Dim Texto As String
    Dim cleanOutput As String
    Dim lastLineWasEmpty As Boolean

    ' Definir los patrones HTML
    htmlPattern = "<(\/?(p|a|ul|b|strong|i|u|br)[^>]*?)>|<\/p><p>"
    liPattern = "<li[^>]*?>"
    cleanHtmlPattern = "<[^>]+>"

    ' Asume que el rango seleccionado es el que quieres modificar
    Set rng = Selection

    For Each celda In rng
        If Not IsEmpty(celda.value) Then
            If Not celda.HasFormula Then ' Ignora celdas con f?rmulas

                ' Reemplazar tabulaciones con espacios
                celda.value = Replace(celda.value, Chr(9), " ")
                ' Eliminar espacios innecesarios
                celda.value = Application.Trim(celda.value)

                ' Reemplazar <li> por saltos de línea
                celda.value = RegExpReemplazar(celda.value, liPattern, vbLf)

                ' Eliminar etiquetas HTML dejando solo texto
                celda.value = RegExpReemplazar(celda.value, cleanHtmlPattern, vbNullString)

                ' Separar el contenido en líneas
                lineas = Split(celda.value, vbLf)
                cleanOutput = ""
                lastLineWasEmpty = False

                ' Remover saltos de línea al inicio del texto
                i = LBound(lineas)
                Do While i <= UBound(lineas) And Trim(lineas(i)) = ""
                    i = i + 1
                Loop

                ' Unificar líneas sin duplicar saltos de línea innecesarios
                For i = i To UBound(lineas)
                    If Trim(lineas(i)) <> "" Then
                        cleanOutput = cleanOutput & lineas(i) & vbLf
                        lastLineWasEmpty = False
                    ElseIf Not lastLineWasEmpty Then
                        cleanOutput = cleanOutput & vbLf
                        lastLineWasEmpty = True
                    End If
                Next i

                ' Eliminar el último salto de línea si existe
                If Len(cleanOutput) > 0 And Right(cleanOutput, 1) = vbLf Then
                    cleanOutput = Left(cleanOutput, Len(cleanOutput) - 1)
                End If

                ' Reemplazar caracteres de control innecesarios
                Texto = ReemplazarEntidadesHtml(cleanOutput)

                ' Asegurar formato limpio sin saltos de línea redundantes
                NuevoTexto = Trim(Texto)

                ' Asignar el texto limpio a la celda
                celda.value = NuevoTexto
            End If
        End If
    Next celda

    MsgBox "Proceso completado: Se han eliminado los saltos de línea al inicio y limpiado el texto.", vbInformation
End Sub



Sub CYB010_AgregarSaltosLineaATextoGuiones()

    Dim celda As Range
    Dim Texto As String
    Dim partes() As String
    Dim i As Integer

    ' Recorre todas las celdas seleccionadas
    For Each celda In Selection
        ' Verifica si la celda tiene texto
        If celda.HasFormula = False Then
            Texto = celda.value
            ' Verifica si hay guiones en el texto
            If InStr(Texto, "-") > 0 Then
                ' Divide el texto usando el guion como delimitador
                partes = Split(Texto, "-")
                
                ' Reconstruye el texto con saltos de línea apropiados
                Texto = partes(0) ' Primer parte (sin cambio)
                
                For i = 1 To UBound(partes)
                    ' Agrega salto de línea solo si no es la primera parte
                    Texto = Texto & vbNewLine & "- " & partes(i)
                Next i
                
                ' Asigna el texto modificado a la celda
                celda.value = Texto
            End If
        End If
    Next celda
    
End Sub


Sub CYB011_BulletsAGuiones()

 Dim celda As Range
    For Each celda In Selection
        If Not IsEmpty(celda.value) Then
            celda.value = Replace(celda.value, "•", "-")
        End If
    Next celda
    MsgBox "Reemplazo completado.", vbInformation, "Proceso Finalizado"
End Sub


Sub CYB011_MantererSoloURLSEnLinea()

    Dim celda As Range
    Dim contenido As String
    Dim resultado As String
    Dim startPos As Integer
    Dim endPos As Integer
    Dim url As String
    Dim commaPos As Integer
    Dim spacePos As Integer
    Dim lineas() As String
    Dim lineasSinVacias() As String
    Dim idx As Integer
    
    ' Iterar sobre cada celda seleccionada
    For Each celda In Selection
        ' Verificar si la celda tiene texto
        If Not IsEmpty(celda.value) Then
            ' Inicializar resultado vacío para cada celda
            resultado = ""
            
            ' Reemplazar comillas dobles por nada
            contenido = Replace(celda.value, """", "")
            
            ' Reemplazar diferentes saltos de línea con vbLf
            contenido = Replace(Replace(Replace(contenido, vbCrLf, vbLf), vbCr, vbLf), vbLf & vbLf, vbLf)
            
            ' Si el contenido comienza con vbLf, quitarlo
            If Left(contenido, 1) = vbLf Then
                contenido = Mid(contenido, 2)
            End If
            
            ' Si el contenido termina con vbLf, quitarlo
            If Right(contenido, 1) = vbLf Then
                contenido = Left(contenido, Len(contenido) - 1)
            End If
            
            ' Dividir el contenido de la celda en un array de líneas
            lineas = Split(contenido, vbLf)
            
            ' Verificar si el array no est? vacío antes de redimensionar
            If UBound(lineas) >= 0 Then
                ' Crear un nuevo array para almacenar las líneas no vacías
                ReDim lineasSinVacias(0 To UBound(lineas))
                idx = 0
                
                ' Iterar sobre cada línea del array
                For i = LBound(lineas) To UBound(lineas)
                    ' Verificar si la línea est? vacía y no agregarla al nuevo array
                    If Trim(lineas(i)) <> "" Then
                        ' Encontrar las URLs dentro de cada línea
                        startPos = InStr(1, lineas(i), "http")
                        Do While startPos > 0
                            ' Encontrar el final de la URL buscando un espacio o coma
                            commaPos = InStr(startPos, lineas(i), ",")
                            spacePos = InStr(startPos, lineas(i), " ")
                            
                            If commaPos > 0 And (commaPos < spacePos Or spacePos = 0) Then
                                endPos = commaPos
                            ElseIf spacePos > 0 Then
                                endPos = spacePos
                            Else
                                endPos = Len(lineas(i)) + 1
                            End If
                            
                            ' Extraer la URL
                            url = Mid(lineas(i), startPos, endPos - startPos)
                            
                            ' Añadir la URL al resultado
                            resultado = resultado & url & vbCrLf
                            
                            ' Buscar la siguiente URL
                            startPos = InStr(startPos + 1, lineas(i), "http")
                        Loop
                    End If
                Next i
                
                ' Eliminar líneas vacías que podrían quedar al final
                If Len(resultado) > 0 Then
                    If Right(resultado, 1) = vbCrLf Then
                        resultado = Left(resultado, Len(resultado) - 1)
                    End If
                End If
                
                ' Asignar el resultado (solo URLs) a la celda
                celda.value = resultado
            End If
        End If
    Next celda
End Sub






Sub CYB012_PingIPs()
    Dim celda As Range
    Dim ip As String
    Dim objShell As Object
    Dim objExec As Object
    Dim resultado As String
    Dim i As Integer
    Dim respuesta As Boolean
    
    ' Crear objeto Shell para ejecutar comandos
    Set objShell = CreateObject("WScript.Shell")
    
    ' Iterar sobre las celdas seleccionadas
    For Each celda In Selection
        ' Obtener la IP de la celda
        ip = Trim(celda.value)
        
        ' Verificar que la celda no est? vacía
        If ip <> "" Then
            respuesta = False ' Inicializar como no respondida
            
            ' Intentar ping hasta 3 veces
            For i = 1 To 3
                ' Ejecutar el ping y capturar la salida
                Set objExec = objShell.Exec("ping -n 1 -w 500 " & ip)
                resultado = objExec.StdOut.ReadAll
                
                ' Si encuentra "TTL=", la IP respondi?
                If InStr(1, resultado, "TTL=", vbTextCompare) > 0 Then
                    respuesta = True
                    Exit For ' Salir del bucle si ya respondi?
                End If
            Next i
            
            ' Cambiar color de celda según el resultado
            If respuesta Then
                celda.Interior.Color = RGB(144, 238, 144) ' Verde claro (IP en línea)
            Else
                celda.Interior.Color = RGB(169, 169, 169) ' Gris oscuro (No responde)
            End If
        End If
    Next celda
    
    ' Liberar objetos
    Set objShell = Nothing
    Set objExec = Nothing
    
    MsgBox "Ping completado.", vbInformation, "Finalizado"
End Sub






Function RegExpReemplazar(ByVal text As String, ByVal replacePattern As String, ByVal replaceWith As String) As String
    ' Funci?n para reemplazar utilizando expresiones regulares
    Dim regex       As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    With regex
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .pattern = replacePattern
    End With
    
    RegExpReemplazar = regex.Replace(text, replaceWith)
End Function

Function ReemplazarEntidadesHtml(ByVal text As String) As String
    ' Funci?n para reemplazar entidades HTML con caracteres correspondientes
    text = Replace(text, "&lt;", "<")
    text = Replace(text, "&gt;", ">")
    text = Replace(text, "&amp;", "&")
    text = Replace(text, "&quot;", """")
    text = Replace(text, "&apos;", "'")        ' Comillas simples
    text = Replace(text, "&#x27;", "'")        ' Comillas simples
    text = Replace(text, "&#34;", """")        ' Comillas dobles
    text = Replace(text, "&#39;", "'")         ' Comillas simples
    text = Replace(text, "&#160;", Chr(160))   ' Espacio no separable
    
    ReemplazarEntidadesHtml = text
End Function

Sub CYB026_OrdenaSegunColorRelleno()
    Dim celdaActual As Range
    Dim tabla As ListObject
    Dim ws As Worksheet
    Dim respuesta As VbMsgBoxResult
    Dim colores As Variant
    Dim i As Integer
    
    ' Definir el orden de los colores (morado ? rojo ? amarillo ? verde)
    colores = Array(RGB(112, 48, 160), RGB(255, 0, 0), RGB(255, 255, 0), RGB(0, 176, 80))
    
    ' Obtener la celda actualmente seleccionada y la hoja activa
    Set celdaActual = ActiveCell
    Set ws = ActiveSheet
    
    ' Verificar si la celda seleccionada est? dentro de una tabla
    On Error Resume Next
    Set tabla = celdaActual.ListObject
    On Error GoTo 0
    
    If tabla Is Nothing Then
        MsgBox "Debes seleccionar una celda dentro de una tabla para ejecutar la ordenaci?n.", vbExclamation, "Error"
        Exit Sub
    End If
    
    ' Confirmar con el usuario antes de proceder
    respuesta = MsgBox("Se ordenar? la tabla por la columna 'Severidad' según el color de relleno." & vbCrLf & _
                       "Orden: Morado ? Rojo ? Amarillo ? Verde." & vbCrLf & vbCrLf & _
                       "¿Deseas continuar?", vbYesNo + vbQuestion, "Confirmaci?n")
    
    If respuesta <> vbYes Then Exit Sub
    
    ' Aplicar ordenaci?n por color en el orden definido
    With ws.ListObjects(tabla.Name).Sort
        .SortFields.Clear
        For i = LBound(colores) To UBound(colores)
            .SortFields.Add key:=tabla.ListColumns("Severidad").Range, _
                            SortOn:=xlSortOnCellColor, _
                            Order:=xlDescending, _
                            DataOption:=xlSortNormal
            .SortFields(1).SortOnValue.Color = colores(i)
            .Apply
        Next i
    End With
    
    MsgBox "Ordenaci?n completada: Morado ? Rojo ? Amarillo ? Verde.", vbInformation, "Proceso finalizado"
End Sub

' GenerarDocumentosWord

Sub EliminarUltimasFilasSiEsSalidaPruebaSeguridad(WordDoc As Object, replaceDic As Object)
    Dim salidaPruebaSeguridadKey As String
    Dim metodoDeteccionKey As String
    Dim salidaPruebaSeguridadValue As String
    Dim metodoDeteccionValue As String
    Dim firstTable As Object
    Dim numRows As Integer
    Dim lastRow As Object
    Dim lastCell As Object
    Dim internalTable As Object

    salidaPruebaSeguridadKey = "«Salidas de herramienta»"
    metodoDeteccionKey = "«Método de detección»"
    
    ' Verificar si las claves est?n presentes en el diccionario
    If replaceDic.Exists(salidaPruebaSeguridadKey) And replaceDic.Exists(metodoDeteccionKey) Then
        ' Obtener los valores de las claves
        salidaPruebaSeguridadValue = CStr(replaceDic(salidaPruebaSeguridadKey))
        metodoDeteccionValue = CStr(replaceDic(metodoDeteccionKey))
        
        ' Inicializar la tabla
        Set firstTable = WordDoc.Tables(1)
        numRows = firstTable.Rows.Count
        
        ' Verificar si ambos valores estan vacíos
        If Len(Trim(salidaPruebaSeguridadValue)) = 0 And Len(Trim(metodoDeteccionValue)) = 0 Then
            ' Si ambos están vacíos, eliminar las últimas filas de la tabla principal
            If numRows > 0 Then
                ' Eliminar la última fila
                firstTable.Rows(numRows).Delete
                ' Eliminar la penúltima fila si hay m?s de una fila
                If numRows > 1 Then
                    firstTable.Rows(numRows - 1).Delete
                End If
            End If
      ElseIf Len(Trim(salidaPruebaSeguridadValue)) = 0 Then
            If numRows > 0 Then
                firstTable.Tables(1).Delete
            End If
      
        End If
    End If
End Sub





Function EstiloExiste(docWord As Object, estilo As String) As Boolean
    Dim st          As Object
    On Error Resume Next
    Set st = docWord.Styles(estilo)
    EstiloExiste = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Sub CYB015_CrearEstilo(docWord As Object, estilo As String)
    Dim nuevoEstilo As Object
    On Error Resume Next
    Set nuevoEstilo = docWord.Styles.Add(Name:=estilo, Type:=1)        ' Tipo 1 = Estilo de p?rrafo
    If Err.Number <> 0 Then
        MsgBox "No se pudo crear el estilo        '" & estilo & "'. " & Err.Description, vbExclamation
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Sub CYB038_ExportarTablaContenidoADocumentoWord()
    Dim appWord     As Object
    Dim docWord     As Object
    Dim tbl         As ListObject
    Dim r           As ListRow
    Dim estilo      As String
    Dim seccion     As String
    Dim descripcion As String
    Dim imagen      As String
    Dim imagenRutaCompleta As String
    Dim parrafo     As Object
    Dim rng         As Object
    Dim rutaBase    As String
    Dim ws          As Worksheet
    Dim parrafoResultados As String
    Dim shape       As Object
    
    ' Obtener la hoja activa
    Set ws = ActiveSheet
    
    ' Obtener la ruta de la carpeta donde est? la hoja activa
    rutaBase = ws.Parent.path & "\"
    
    ' Inicializar Word
    Set appWord = CreateObject("Word.Application")
    appWord.Visible = True
    Set docWord = appWord.Documents.Add
    
    ' Definir la tabla activa de Excel
    On Error Resume Next
    Set tbl = ws.ListObjects("Tabla_pruebas_seguridad")
    On Error GoTo 0
    
    If tbl Is Nothing Then
        MsgBox "La tabla        'Tabla_pruebas_seguridad' no se encuentra en la hoja activa.", vbExclamation
        GoTo CleanUp
    End If
    
    ' Procesar cada fila de la tabla
    For Each r In tbl.ListRows
        estilo = r.Range.Cells(1, tbl.ListColumns("Estilo").Index).value
        seccion = r.Range.Cells(1, tbl.ListColumns("Secci?n").Index).value
        descripcion = r.Range.Cells(1, tbl.ListColumns("Descripci?n").Index).value
        imagen = r.Range.Cells(1, tbl.ListColumns("Im?genes").Index).value
        parrafoResultados = r.Range.Cells(1, tbl.ListColumns("Resultado").Index).value
        
        ' Obtener la ruta completa de la imagen
        imagenRutaCompleta = rutaBase & imagen
        
        ' Verificar y crear el estilo si no existe
        If Not EstiloExiste(docWord, estilo) Then
            CrearEstilo docWord, estilo
        End If
        
        ' Agregar el encabezado con el estilo correspondiente
        If Trim(seccion) <> "" Then
            With docWord.content.Paragraphs.Add
                .Range.text = seccion
                .Range.Style = docWord.Styles(estilo)
                .Range.InsertParagraphAfter
            End With
            
        End If
        
        ' Agregar un p?rrafo con la descripci?n
        If Trim(descripcion) <> "" Then
            With docWord.content.Paragraphs.Add
                .Range.text = descripcion
                .Range.Style = docWord.Styles("Normal")        ' Aplicar un estilo predeterminado para el p?rrafo de descripci?n
                .Format.SpaceBefore = 12        ' Espacio antes del p?rrafo para separaci?n
            End With
        End If
        
        docWord.content.InsertParagraphAfter
        docWord.content.Paragraphs.Last.Range.Select
        
        ' Agregar la imagen si existe
        If imagen <> "" Then
            ' Verificar si la imagen existe
            If Dir(imagenRutaCompleta) <> "" Then
                ' Agregar un p?rrafo vacío para la imagen
                docWord.content.InsertParagraphAfter
                Set rng = docWord.content.Paragraphs.Last.Range
                
                ' Insertar la imagen en el p?rrafo vacío
                Set shape = docWord.InlineShapes.AddPicture(fileName:=imagenRutaCompleta, LinkToFile:=False, SaveWithDocument:=True, Range:=rng)
                
                ' Centrar la imagen
                shape.Range.ParagraphFormat.Alignment = 1        ' 1 = wdAlignParagraphCenter
                
                ' Agregar un p?rrafo vacío despu?s de la imagen para el caption
                docWord.content.InsertParagraphAfter
                Set rng = docWord.content.Paragraphs.Last.Range
                
                ' Insertar el caption debajo de la imagen
                Set captionRange = rng.Duplicate
                captionRange.Select
                appWord.Selection.MoveLeft unit:=1, Count:=1, Extend:=0        ' wdCharacter
                appWord.CaptionLabels.Add Name:="Imagen"
                appWord.Selection.InsertCaption Label:="Imagen", TitleAutoText:="InsertarTítulo1", _
                                                Title:="", Position:=1        ' wdCaptionPositionBelow, ExcludeLabel:=0
                appWord.Selection.ParagraphFormat.Alignment = 1        ' wdAlignParagraphCenter
                
                docWord.content.InsertAfter text:=" " & seccion
                
                ' Agregar un p?rrafo vacío despu?s del caption para separaci?n
                docWord.content.InsertParagraphAfter
                
            Else
                MsgBox "La imagen        '" & imagenRutaCompleta & "' no se encuentra.", vbExclamation
            End If
        End If
        
        ' Agregar el p?rrafo de resultados si no est? vacío
        If Trim(parrafoResultados) <> "" Then
            With docWord.content.Paragraphs.Add
                .Range.text = parrafoResultados
                .Range.Style = docWord.Styles("Normal")        ' Aplicar un estilo predeterminado para el p?rrafo de resultados
                .Format.SpaceBefore = 12        ' Espacio antes del p?rrafo para separaci?n
            End With
        End If
    Next r
    
CleanUp:
    ' Limpiar
    Set appWord = Nothing
    Set docWord = Nothing
    Set tbl = Nothing
    Set rng = Nothing
    Set shape = Nothing
End Sub

Sub CYB017_LimpiarColumnaReferencias()
    Dim rng         As Range
    Dim cell        As Range
    Dim content     As String
    Dim contentArray() As String
    Dim i           As Integer
    Dim temp        As String
    Dim uniqueUrls  As Object
    Dim uniqueArray() As String
    Dim n           As Integer
    Dim newContent  As String
    Dim filteredContent As String
    
    ' Selecciona el rango deseado
    Set rng = Selection
    
    ' Recorre cada celda en el rango
    For Each cell In rng
        ' Obtiene el contenido de la celda
        content = cell.value
        
        ' Sustituye comillas dobles con saltos de línea (Char 10)
        content = Replace(content, """", Chr(10))
        
        ' Comprueba si el contenido es vacío
        If content <> "" Then
            ' Convierte el contenido en un array separado por el car?cter de nueva línea
            contentArray = Split(content, Chr(10))
            
            ' Inicializa el diccionario para almacenar las URL únicas
            Set uniqueUrls = CreateObject("Scripting.Dictionary")
            
            ' Agrega las URL únicas al diccionario
            For i = LBound(contentArray) To UBound(contentArray)
                If Trim(contentArray(i)) <> "" Then
                    ' Elimina espacios en blanco, Chr(10) y Chr(13) del elemento
                    contentArray(i) = Trim(Replace(contentArray(i), Chr(13), ""))
                    contentArray(i) = Replace(contentArray(i), " ", "")
                    If InStr(1, contentArray(i), "wikipedia", vbTextCompare) = 0 Then
                        If Not uniqueUrls.Exists(contentArray(i)) Then
                            uniqueUrls.Add contentArray(i), Nothing
                        End If
                    End If
                End If
            Next i
            
            ' Convertir la colecci?n de claves en un array
            n = uniqueUrls.Count - 1
            ReDim uniqueArray(n)
            i = 0
            For Each key In uniqueUrls.Keys
                uniqueArray(i) = key
                i = i + 1
            Next
            
            ' Ordena los elementos del array
            For i = LBound(uniqueArray) To UBound(uniqueArray) - 1
                For j = i + 1 To UBound(uniqueArray)
                    If uniqueArray(i) > uniqueArray(j) Then
                        temp = uniqueArray(i)
                        uniqueArray(i) = uniqueArray(j)
                        uniqueArray(j) = temp
                    End If
                Next j
            Next i
            
            ' Convierte el array nuevamente en una cadena concatenada por el car?cter de nueva línea
            newContent = Join(uniqueArray, Chr(10))
            
            ' Elimina saltos de línea iniciales y finales
            newContent = Trim(newContent)
            
            ' Filtra líneas para conservar solo las que contienen "//"
            contentArray = Split(newContent, Chr(10))
            filteredContent = ""
            
            For i = LBound(contentArray) To UBound(contentArray)
                If InStr(contentArray(i), "//") > 0 Then
                    If filteredContent <> "" Then
                        filteredContent = filteredContent & Chr(10)
                    End If
                    filteredContent = filteredContent & contentArray(i)
                End If
            Next i
            
            ' Asigna el contenido filtrado a la celda
            cell.value = filteredContent
        End If
    Next cell
End Sub

Sub CYB018_LeerArchivoTXT(txtFilePath As String, dataDict As Object)
    Dim fileNumber  As Integer
    Dim line        As String
    Dim keyValue()  As String
    Dim key         As String
    Dim value       As String
    
    fileNumber = FreeFile
    Open txtFilePath For Input As #fileNumber
    
    Do While Not EOF(fileNumber)
        Line Input #fileNumber, line
        ' Divide la línea en clave y valor
        keyValue = Split(line, ":")
        If UBound(keyValue) = 1 Then
            key = Trim(keyValue(0))
            value = Trim(Mid(keyValue(1), 2, Len(keyValue(1)) - 2))        ' Extrae el valor entre comillas dobles
            ' Añadir al diccionario
            dataDict(key) = value
        End If
    Loop
    
    Close #fileNumber
End Sub

Sub CYB019_WordAppAlternativeReplaceParagraph(WordApp As Object, WordDoc As Object, wordToFind As String, replaceWord As String)
    Dim findInRange As Boolean
    Dim rng         As Object
    
    ' Establecer el rango al contenido del documento
    Set rng = WordDoc.content
    
    ' Configurar la búsqueda
    With rng.Find
        .text = wordToFind
        .Replacement.text = replaceWord
        .Forward = True
        .Wrap = 1        ' wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    
    ' Reemplazar todo el texto encontrado
    rng.Find.Execute Replace:=2        ' wdReplaceAll
End Sub


Function ActualizarGraficoSegunDicionario(ByRef WordDoc As Object, conteos As Object, graficoIndex As Integer) As Boolean
    Dim ils As Object
    Dim Chart As Object
    Dim ChartData As Object
    Dim ChartWorkbook As Object
    Dim SourceSheet As Object
    Dim dataRangeAddress As String
    Dim categoryRow As Integer
    Dim category As Variant
    Dim lastRow As Long
    Dim sheetIndex As Integer
    
    On Error GoTo ErrorHandler
    
    ' Verificar que el índice del gr?fico es v?lido
    If graficoIndex < 1 Or graficoIndex > WordDoc.InlineShapes.Count Then
        MsgBox "Índice de gr?fico fuera de rango."
        ActualizarGraficoSegunDicionario = False
        Exit Function
    End If
    
    ' Obtener el InlineShape correspondiente al índice
    Set ils = WordDoc.InlineShapes(graficoIndex)
    
    If ils.Type = 12 And ils.HasChart Then
        Set Chart = ils.Chart
        If Not Chart Is Nothing Then
            ' Activar el libro de trabajo asociado al gr?fico
            Set ChartData = Chart.ChartData
            If Not ChartData Is Nothing Then
                ChartData.Activate
                Set ChartWorkbook = ChartData.Workbook
                If Not ChartWorkbook Is Nothing Then
                    Set SourceSheet = ChartWorkbook.Sheets(1)        ' Usar la primera hoja del libro
                    
                    ' Limpiar las celdas en el rango antes de agregar nuevos datos
                    lastRow = SourceSheet.Cells(SourceSheet.Rows.Count, 1).End(xlUp).Row
                    If lastRow >= 2 Then
                        SourceSheet.Range("A2:B" & lastRow).ClearContents
                    End If
                    
                    ' Insertar nuevos datos
                    categoryRow = 2        ' Empezar en la fila 2 para los tipos de vulnerabilidad
                    For Each category In conteos.Keys
                        SourceSheet.Cells(categoryRow, 1).value = category
                        SourceSheet.Cells(categoryRow, 2).value = conteos(category)
                        categoryRow = categoryRow + 1
                    Next category
                    
                    ' Construir el rango din?mico como una cadena
                    sheetIndex = 1
                    dataRangeAddress = CStr(ChartWorkbook.Sheets(sheetIndex).Name & "$A$1:$B$" & CStr(categoryRow - 1))
                    Debug.Print dataRangeAddress
                    
                    ' Actualizar el gr?fico con el nuevo rango de datos
                    On Error Resume Next
                    ChartWorkbook.Sheets(sheetIndex).ChartObjects(1).Chart.SetSourceData Source:=Range(dataRangeAddress)
                    If Err.Number <> 0 Then
                        MsgBox "Error al establecer el rango de datos: " & Err.Description
                        Err.Clear
                        ActualizarGraficoSegunDicionario = False
                        Exit Function
                    End If
                    On Error GoTo 0
                    
                    ' Actualizar el gr?fico
                    On Error Resume Next
                    Chart.Refresh
                    If Err.Number <> 0 Then
                        MsgBox "Error al actualizar el gr?fico: " & Err.Description
                        Err.Clear
                        ActualizarGraficoSegunDicionario = False
                        Exit Function
                    End If
                    On Error GoTo 0
                    
                    ' Cerrar el libro de trabajo sin guardar cambios
                    ChartWorkbook.Close SaveChanges:=False
                    
                    ActualizarGraficoSegunDicionario = True
                End If
            End If
        Else
            MsgBox "El InlineShape seleccionado no contiene un gr?fico v?lido."
            ActualizarGraficoSegunDicionario = False
        End If
    Else
        MsgBox "El InlineShape seleccionado no contiene un gr?fico."
        ActualizarGraficoSegunDicionario = False
    End If
    
    Exit Function
    
ErrorHandler:
    MsgBox "Ocurri? un error: " & Err.Description, vbCritical
    ActualizarGraficoSegunDicionario = False
End Function

Sub CYB001_GenerarDocumentosVulnerabilidadesWord()
    Dim rng As Range
    Dim tbl As ListObject
    Dim WordApp As Object 'Word.Application ' Requiere referencia a Microsoft Word Object Library
    Dim WordDoc As Object 'Word.Document ' Requiere referencia a Microsoft Word Object Library
    Dim templatePath As String
    Dim outputPath As String
    Dim replaceDic As Object 'Scripting.Dictionary
    Dim cell As Range
    Dim headerCell As Range
    Dim colIndex As Integer
    Dim explicacionTecnicaCol As Long
    Dim tipoTextoExplicacionCol As Long
    Dim rowCount As Long ' Usar Long para números de fila
    Dim i As Long      ' Usar Long para números de fila
    Dim tempFolder As String
    Dim tempFolderPath As String
    Dim saveFolder As String
    Dim selectedRange As Range ' Variable para almacenar el rango seleccionado por el usuario
    Dim documentsList() As String ' Lista para almacenar los documentos generados
    Dim fs As Object 'Scripting.FileSystemObject
    Dim key As Variant
    Dim explicacionTecnicaValue As String
    Dim tipoTextoValue As String
    Dim excelBasePath As String
    Dim finalDocumentPath As String ' Ruta para guardar el documento final
    Dim textoCelda As String
    
    ' --- Selección de Rango y Verificación de Tabla ---
    On Error Resume Next
    Set selectedRange = Application.InputBox("Seleccione TODO el rango de la tabla (incluyendo encabezados) que contiene los datos", Type:=8)
    On Error GoTo 0
    
    If selectedRange Is Nothing Then
        MsgBox "No se ha seleccionado un rango válido.", vbExclamation
        Exit Sub
    End If
    
    On Error Resume Next
    Set tbl = selectedRange.ListObject
    On Error GoTo 0
    
    If tbl Is Nothing Then
        MsgBox "El rango seleccionado no está dentro de una tabla (ListObject)." & vbCrLf & _
               "Asegúrese de que los datos estén formateados como tabla (Insertar > Tabla).", vbExclamation
        Exit Sub
    End If
    
    Set rng = tbl.Range ' Usar el rango completo de la tabla
    
    ' --- Selección de Plantilla y Carpeta de Salida ---
    templatePath = Application.GetOpenFilename("Documentos de Word (*.docx), *.docx", , "Seleccione un documento de Word como plantilla")
    If templatePath = "Falso" Then Exit Sub
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Seleccione la carpeta donde desea guardar los documentos generados"
        If .Show = -1 Then
            saveFolder = .SelectedItems(1)
            If Right(saveFolder, 1) <> "\" Then saveFolder = saveFolder & "\" ' Asegurar que termina con \
        Else
            Exit Sub ' Usuario canceló
        End If
    End With
    
    ' --- Inicializar Word ---
    On Error Resume Next
    Set WordApp = GetObject(, "Word.Application") ' Intenta conectar con instancia existente
    If Err.Number <> 0 Then
        Set WordApp = CreateObject("Word.Application") ' Crea nueva instancia si no existe
        Err.Clear
    End If
    On Error GoTo 0 ' Restablecer manejo de errores
    
    If WordApp Is Nothing Then
        MsgBox "No se pudo iniciar Microsoft Word.", vbCritical
        Exit Sub
    End If
    WordApp.Visible = True
    
    ' --- Preparar Carpeta Temporal y FileSystemObject ---
    Set fs = CreateObject("Scripting.FileSystemObject")
    tempFolder = Environ("TEMP") & "\tmp-" & Format(Now(), "yyyymmddhhmmss")
    If Not fs.FolderExists(tempFolder) Then MkDir tempFolder
    tempFolderPath = tempFolder & "\"
    
    ' Ruta base para imágenes relativas (directorio del archivo Excel)
    If ThisWorkbook.path <> "" Then
        excelBasePath = ThisWorkbook.path & "\"
    Else
        MsgBox "Guarde primero el libro de Excel para poder resolver rutas relativas de imágenes.", vbExclamation
        ' Opcionalmente, pedir una ruta base al usuario
        ' excelBasePath = InputBox("Ingrese la ruta base para las imágenes relativas:")
        ' If Right(excelBasePath, 1) <> "\" Then excelBasePath = excelBasePath & "\"
        WordApp.Quit
        Set WordApp = Nothing
        Set fs = Nothing
        Exit Sub
    End If

    ' --- Encontrar Índices de Columnas Clave (más robusto) ---
    explicacionTecnicaCol = 0
    tipoTextoExplicacionCol = 0
    For Each headerCell In tbl.HeaderRowRange.Cells
        Select Case Trim(headerCell.value)
            Case "Explicación técnica"
                explicacionTecnicaCol = headerCell.Column - tbl.Range.Column + 1
            Case "Tipo de texto de explicación técnica"
                tipoTextoExplicacionCol = headerCell.Column - tbl.Range.Column + 1
        End Select
    Next headerCell
    
    If explicacionTecnicaCol = 0 Then
        MsgBox "No se encontró la columna 'Explicación técnica' en la tabla.", vbCritical
        WordApp.Quit
        Set WordApp = Nothing
        Set fs = Nothing
        Exit Sub
    End If
    ' Permitir que 'Tipo de texto de explicación técnica' sea opcional
    If tipoTextoExplicacionCol = 0 Then
        MsgBox "Advertencia: No se encontró la columna 'Tipo de texto de explicación técnica'." & vbCrLf & _
               "Se asumirá Texto Plano para todas las filas.", vbInformation
    End If

    ' --- Bucle Principal para Generar Documentos ---
    rowCount = tbl.ListRows.Count
    ReDim documentsList(0 To rowCount - 1) ' Ajustar tamaño del array
    
    For i = 1 To rowCount ' Iterar sobre las filas de datos de la tabla
        ' Crear un nuevo diccionario para cada fila
        Set replaceDic = CreateObject("Scripting.Dictionary")
        
        ' Llena el diccionario con los datos de la fila actual
        For colIndex = 1 To tbl.ListColumns.Count
            Dim colName As String
            Dim cellValue As String
            colName = tbl.HeaderRowRange.Cells(1, colIndex).value
            cellValue = tbl.DataBodyRange.Cells(i, colIndex).value
            replaceDic("«" & colName & "»") = cellValue
            
            ' Guardar valores específicos para manejo especial
            If colIndex = explicacionTecnicaCol Then
                explicacionTecnicaValue = cellValue
            End If
            If tipoTextoExplicacionCol > 0 And colIndex = tipoTextoExplicacionCol Then
                tipoTextoValue = Trim(LCase(cellValue)) ' Convertir a minúsculas y quitar espacios
            ElseIf tipoTextoExplicacionCol = 0 Then
                 tipoTextoValue = "texto plano" ' Valor por defecto si la columna no existe
            End If
        Next colIndex
        
        ' Crear y abrir copia del documento temporal
        Dim tempDocPath As String
        tempDocPath = tempFolderPath & "Documento_" & i & ".docx"
        fs.CopyFile templatePath, tempDocPath, True ' True para sobrescribir si existe
        Set WordDoc = WordApp.Documents.Open(tempDocPath)
        WordDoc.Activate
        
        ' Realiza los reemplazos en el documento de Word
        For Each key In replaceDic.Keys
            Dim placeholder As String
            Dim replacementValue As String
            placeholder = CStr(key)
            replacementValue = CStr(replaceDic(key))
            
            If placeholder = "«Explicación técnica»" Then
                If tipoTextoValue = "markdown" Then
                    'RawPrint explicacionTecnicaValue
                    'explicacionTecnicaValue = Replace(explicacionTecnicaValue, vbLf & vbLf, vbLf)
                    'RawPrint explicacionTecnicaValue
                    InsertarTextoMarkdownEnWordConFormato WordApp, WordDoc, placeholder, explicacionTecnicaValue, excelBasePath
                Else
                    WordAppReemplazarParrafo WordApp, WordDoc, placeholder, replacementValue
                End If
            ElseIf placeholder = "«Descripción»" Then
                 WordAppReemplazarParrafo WordApp, WordDoc, placeholder, replacementValue ' O usa reemplazo directo si aplica
            ElseIf placeholder = "«Propuesta de remediación»" Then
                 WordAppReemplazarParrafo WordApp, WordDoc, placeholder, replacementValue
            ElseIf placeholder = "«Referencias»" Then
                 WordAppReemplazarParrafo WordApp, WordDoc, placeholder, replacementValue
            Else
                WordAppReemplazarParrafo WordApp, WordDoc, placeholder, replacementValue
            End If
        Next key
        
        With WordDoc.Tables(1).cell(8, 1).Range
    .Font.Color = wdColorBlack
    .ParagraphFormat.Alignment = wdAlignParagraphJustify
End With
        FormatearCeldaNivelRiesgo WordDoc.Tables(1).cell(1, 2)
       
    ' Condición 1: si la celda (3,1) dice AMENAZA
    textoCelda = Trim(Replace(WordDoc.Tables(1).cell(3, 1).Range.text, Chr(13) & Chr(7), ""))
    If textoCelda = "AMENAZA" Then
        FormatearParrafosGuionesCelda WordDoc.Tables(1).cell(3, 2)
    End If

    ' Condición 2: si la celda (4,1) dice PROPUESTA DE REMEDIACIÓN
    textoCelda = Trim(Replace(WordDoc.Tables(1).cell(4, 1).Range.text, Chr(13) & Chr(7), ""))
    If textoCelda = "PROPUESTA DE REMEDIACIÓN" Then
        FormatearParrafosGuionesCelda WordDoc.Tables(1).cell(4, 2)
    End If

    ' Condición 3: si la celda (5,1) dice AMENAZA o PROPUESTA DE REMEDIACIÓN
    textoCelda = Trim(Replace(WordDoc.Tables(1).cell(5, 1).Range.text, Chr(13) & Chr(7), ""))
    If textoCelda = "AMENAZA" Or textoCelda = "PROPUESTA DE REMEDIACIÓN" Then
        FormatearParrafosGuionesCelda WordDoc.Tables(1).cell(5, 2)
    End If
       
        
        EliminarUltimasFilasSiEsSalidaPruebaSeguridad WordDoc, replaceDic
                
        SustituirTextoMarkdownPorImagenes WordApp, WordDoc, excelBasePath
        
    textoCelda = Trim(Replace(WordDoc.Tables(1).cell(7, 1).Range.text, Chr(13) & Chr(7), ""))
If textoCelda = "DETALLE DE PRUEBAS DE SEGURIDAD" Then
    EliminarLineasVaciasEnCelda WordDoc
End If

        ' Guardar y cerrar el documento individual
        finalDocumentPath = saveFolder & "Documento_Final_" & i & ".docx"
        WordDoc.SaveAs finalDocumentPath
        WordDoc.Close
        
        ' Agregar la ruta del documento generado a la lista
        documentsList(i - 1) = finalDocumentPath
    Next i
    
    ' --- Llamar a la función FusionarDocumentosInsertando ---
    FusionarDocumentosInsertando WordApp, documentsList, saveFolder & "Documento_Completo_Fusionado.docx"
    
    ' Cerrar la aplicaci?n de Word
    WordApp.Quit
    Set WordApp = Nothing
    
    ' Muestra un mensaje de ?xito
    MsgBox "Se han generado los documentos de Word correctamente.", vbInformation
    
    ' Limpiar objetos
    Set replaceDic = Nothing
    Set WordApp = Nothing
    Set fs = Nothing
End Sub


Sub CYB007_GenerarReportesVulnsAppsINAI()
    Dim WordApp     As Object
    Dim WordDoc     As Object
    Dim plantillaReportePath As String
    Dim plantillaReportePath2 As String
    Dim plantillaVulnerabilidadesPath As String
    Dim carpetaSalida As String
    Dim archivoTemp As String
    Dim fileSystem  As Object
    Dim dlg         As Object
    Dim rng         As Object
    Dim replaceDic  As Object
    Dim cell        As Object
    Dim rowCount    As Integer
    Dim i           As Integer
    Dim tempFolder  As String
    Dim tempFolderGenerados As String
    Dim finalDocumentPath As String
    Dim tempDocPath As String
    Dim tempDocVulnerabilidadesPath As String
    Dim tempFileName As String
    Dim secVulnerabilidades As String
    Dim rngReplace  As Object
    Dim campoArchivoPath As String
    Dim partes()    As String
    Dim appName    As String
    Dim key         As String
    Dim value       As String
    Dim documentsList() As String
    Dim numDocuments As Integer
    Dim severityCounts As Object
    Dim severity    As Variant
    Dim totalVulnerabilidades As Integer
    Dim Chart       As Object
    Dim seriesCollection As Object
    Dim dataLabels  As Object
    Dim severidadColumna As Integer
    Dim countBAJA   As Integer
    Dim countMEDIA  As Integer
    Dim countALTA   As Integer
    Dim countCRITICAS As Integer
    Dim vulntypesCounts As Object
    Dim tiposvulnerabilidadColumna As Integer
    Dim vulntypes   As Variant
    Dim ws          As Object
    Dim wb          As Object
    Dim selectedRange As Object
    Dim ChartData   As Object
    Dim ChartWorkbook As Object
    
    ' Crear un di?logo para seleccionar el archivo CSV
    campoArchivoPath = Application.GetOpenFilename("Archivos CSV (*.csv), *.csv", , "Seleccionar archivo CSV")
    If campoArchivoPath = "Falso" Then
        MsgBox "No se seleccion? ningún archivo CSV. La macro se detendr?."
        Exit Sub
    End If
    
    ' Leer campos de reemplazo desde el archivo CSV
    Set replaceDic = CreateObject("Scripting.Dictionary")
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set ts = fileSystem.OpenTextFile(campoArchivoPath, 1, False, 0)
    
    ' Leer el archivo línea por línea
    Do Until ts.AtEndOfStream
        csvLine = ts.ReadLine
        partes = Split(csvLine, ",", 2)        ' Divide en dos partes (clave, valor)
        
        If UBound(partes) = 1 Then
            key = Trim(partes(0))
            value = Trim(partes(1))
            
            ' Añadir al diccionario
            replaceDic(key) = value
        End If
    Loop
    ts.Close
    
    ' Extraer el nombre de la Aplicaci?n del diccionario
    If replaceDic.Exists("«Aplicaci?n»") Then
        appName = replaceDic("«Aplicaci?n»")
    Else
        MsgBox "No se encontr? el campo        'Aplicaci?n' en el archivo CSV.", vbExclamation
        Exit Sub
    End If
    
    ' Crear di?logos para seleccionar plantillas y carpeta de salida
    Set dlg = Application.FileDialog(msoFileDialogFilePicker)
    
    dlg.Title = "Seleccionar la plantilla de reporte t?cnico"
    dlg.Filters.Clear
    dlg.Filters.Add "Archivos de Word", "*.docx"
    If dlg.Show = -1 Then
        plantillaReportePath = dlg.SelectedItems(1)
    Else
        MsgBox "No se seleccion? ningún archivo. La macro se detendr?."
        Exit Sub
    End If
    
    dlg.Title = "Seleccionar la plantilla de reporte ejecutivo"
    If dlg.Show = -1 Then
        plantillaReportePath2 = dlg.SelectedItems(1)
    Else
        MsgBox "No se seleccion? ningún archivo. La macro se detendr?."
        Exit Sub
    End If
    
    dlg.Title = "Seleccionar la plantilla de tabla de vulnerabilidades"
    If dlg.Show = -1 Then
        plantillaVulnerabilidadesPath = dlg.SelectedItems(1)
    Else
        MsgBox "No se seleccion? ningún archivo. La macro se detendr?."
        Exit Sub
    End If
    
    ' Solicitar la carpeta de salida
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Seleccionar Carpeta de Salida"
        If .Show = -1 Then
            carpetaSalida = .SelectedItems(1)
        Else
            MsgBox "No se seleccion? ninguna carpeta. La macro se detendr?."
            Exit Sub
        End If
    End With
    
    ' Crear una subcarpeta con el nombre de la Aplicaci?n
    carpetaSalida = carpetaSalida & "\AV " & appName
    On Error Resume Next
    MkDir carpetaSalida
    On Error GoTo 0
    
    ' Crear una instancia de Word
    On Error Resume Next
    Set WordApp = CreateObject("Word.Application")
    On Error GoTo 0
    
    If WordApp Is Nothing Then
        MsgBox "No se puede iniciar Microsoft Word."
        Exit Sub
    End If
    
    ' Crear una carpeta temporal
    tempFolder = Environ("TEMP") & "\tmp-" & Format(Now(), "yyyymmddhhmmss")
    On Error Resume Next
    MkDir tempFolder
    On Error GoTo 0
    
    ' Crear una carpeta para documentos generados
    tempFolderGenerados = tempFolder & "\Documentos generados"
    On Error Resume Next
    MkDir tempFolderGenerados
    On Error GoTo 0
    
    ' Crear y abrir documentos de reporte t?cnico y ejecutivo
    For Each plantilla In Array(plantillaReportePath, plantillaReportePath2)
        archivoTemp = tempFolder & "\" & fileSystem.GetFileName(plantilla)
        fileSystem.CopyFile plantilla, archivoTemp
        
        Set WordDoc = WordApp.Documents.Open(archivoTemp)
        WordApp.Visible = False
        
        ReemplazarCampos WordDoc, replaceDic
        
        If plantilla = plantillaReportePath Then
            tempDocPath = tempFolder & "\SSIFO14-03 Informe t?cnico.docx"
            WordDoc.SaveAs2 tempDocPath
        Else
            tempDocPath2 = tempFolder & "\SSIFO15-03 Informe Ejecutivo.docx"
            WordDoc.SaveAs2 tempDocPath2
        End If
        
        WordDoc.Close False
    Next
    
    ' Solicita al usuario seleccionar el rango de celdas
    On Error Resume Next
    Set selectedRange = Application.InputBox("Seleccione el rango de celdas que contienen las columnas a considerar", Type:=8)
    On Error GoTo 0
    
    If selectedRange Is Nothing Then
        MsgBox "No se ha seleccionado un rango v?lido.", vbExclamation
        Exit Sub
    End If
    
    Dim resultado As Boolean
    
    ' Llamar a la funci?n para exportar la hoja activa a Excel
    resultado = FunExportarHojaActivaAExcelINAI(carpetaSalida, appName)
    
    ' Verifica si el rango seleccionado est? dentro de una tabla
    On Error Resume Next
    Set rng = selectedRange.ListObject.Range
    On Error GoTo 0
    
    If rng Is Nothing Then
        MsgBox "El rango seleccionado no est? dentro de una tabla.", vbExclamation
        Exit Sub
    End If
    
    ' Inicializar el diccionario para contar severidades
    Set severityCounts = CreateObject("Scripting.Dictionary")
    
    ' Buscar la columna "Severidad"
    severidadColumna = -1
    For i = 1 To rng.Columns.Count
        If rng.Cells(1, i).value = "Severidad" Then
            severidadColumna = i
            Exit For
        End If
    Next i
    
    If severidadColumna = -1 Then
        MsgBox "No se encontr? la columna        'Severidad' en el rango seleccionado.", vbExclamation
        Exit Sub
    End If
    
    ' Contar severidades
    For i = 2 To rng.Rows.Count
        severity = rng.Cells(i, severidadColumna).value
        If severity <> "" Then
            If severityCounts.Exists(severity) Then
                severityCounts(severity) = severityCounts(severity) + 1
            Else
                severityCounts.Add severity, 1
            End If
        End If
    Next i
    
    ' Inicializar el diccionario para contar tipos de vulnerabilidades
    Set vulntypesCounts = CreateObject("Scripting.Dictionary")
    
    ' Buscar la columna "Tipo de vulnerabilidad"
    tiposvulnerabilidadColumna = -1
    For i = 1 To rng.Columns.Count
        If rng.Cells(1, i).value = "Tipo de vulnerabilidad" Then
            tiposvulnerabilidadColumna = i
            Exit For
        End If
    Next i
    
    If tiposvulnerabilidadColumna = -1 Then
        MsgBox "No se encontr? la columna        'Tipo de vulnerabilidad' en el rango seleccionado.", vbExclamation
        Exit Sub
    End If
    
    ' Contar tipos de vulnerabilidades
    For i = 2 To rng.Rows.Count
        vulntypes = rng.Cells(i, tiposvulnerabilidadColumna).value
        If vulntypes <> "" Then
            If vulntypesCounts.Exists(vulntypes) Then
                vulntypesCounts(vulntypes) = vulntypesCounts(vulntypes) + 1
            Else
                vulntypesCounts.Add vulntypes, 1
            End If
        End If
    Next i
    
    ' Inicializar conteos
    countBAJA = IIf(severityCounts.Exists("BAJA"), severityCounts("BAJA"), 0)
    countMEDIA = IIf(severityCounts.Exists("MEDIA"), severityCounts("MEDIA"), 0)
    countALTA = IIf(severityCounts.Exists("ALTA"), severityCounts("ALTA"), 0)
    countCRITICAS = IIf(severityCounts.Exists("CRÍTICOS"), severityCounts("CRÍTICOS"), 0)
    
    ' Calcular total de vulnerabilidades
    totalVulnerabilidades = countBAJA + countMEDIA + countALTA + countCRITICAS
    
    ' Copia la plantilla de vulnerabilidades
    tempDocVulnerabilidadesPath = tempFolder & "\Plantilla_Vulnerabilidades.docx"
    fileSystem.CopyFile plantillaVulnerabilidadesPath, tempDocVulnerabilidadesPath
    
    ' Genera documentos de Word por cada registro
    rowCount = rng.Rows.Count
    For i = 2 To rowCount
        Set replaceDic = CreateObject("Scripting.Dictionary")
        
        For Each cell In selectedRange.Rows(1).Cells
            replaceDic("«" & cell.value & "»") = rng.Cells(i, cell.Column).value
        Next cell
        
        tempFileName = "Documento_" & i & ".docx"
        fileSystem.CopyFile tempDocVulnerabilidadesPath, tempFolderGenerados & "\" & tempFileName
        
        On Error Resume Next
        Set WordDoc = WordApp.Documents.Open(tempFolderGenerados & "\" & tempFileName)
        If Err.Number <> 0 Then
            MsgBox "No se pudo abrir el archivo: " & tempFolderGenerados & "\" & tempFileName, vbCritical
            Err.Clear
            Exit Sub
        End If
        On Error GoTo 0
        
        ReemplazarCampos WordDoc, replaceDic
        FormatearCeldaNivelRiesgo WordDoc.Tables(1).cell(1, 2)
        EliminarUltimasFilasSiEsSalidaPruebaSeguridad WordDoc, replaceDic
        WordDoc.Save
        WordDoc.Close
        
        numDocuments = numDocuments + 1
        ReDim Preserve documentsList(numDocuments - 1)
        documentsList(numDocuments - 1) = tempFolderGenerados & "\" & tempFileName
    Next i
    
    ' Combina todos los archivos en uno solo
    finalDocumentPath = tempFolder & "\Tablas_vulnerabilidades.docx"
    FusionarDocumentosInsertando WordApp, documentsList, finalDocumentPath
    
    ' Actualizar el documento de reporte t?cnico
    Set WordDoc = WordApp.Documents.Open(tempDocPath)
    secVulnerabilidades = "{{Secci?n de tablas de vulnerabilidades}}"
    
    Set rngReplace = WordDoc.content
    rngReplace.Find.ClearFormatting
    With rngReplace.Find
        .text = secVulnerabilidades
        .Replacement.text = ""
        .Forward = True
        .Wrap = 1        ' wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    
    If rngReplace.Find.Execute Then
        rngReplace.text = ""
        rngReplace.InsertFile finalDocumentPath
    End If
    
    ' Reemplazar el total de vulnerabilidades
    Set rngReplace = WordDoc.content
    rngReplace.Find.ClearFormatting
    With rngReplace.Find
        .text = "«Total de vulnerabilidades»"
        .Replacement.text = totalVulnerabilidades
        .Forward = True
        .Wrap = 1        ' wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    
    If rngReplace.Find.Execute Then
        rngReplace.text = totalVulnerabilidades
    End If
    
    ' Actualizar el gr?fico InlineShape n?mero 1 en reporte t?cnico
    FunActualizarGraficoSegunDicionario WordDoc, severityCounts, 1
    
    ' Actualizar todos los gr?ficos en el documento
    ActualizarGraficos WordDoc
    ' Update the Table of Contents
    On Error Resume Next
    WordDoc.TablesOfContents(1).Update
    On Error GoTo 0
    
    ' Guardar el documento de reporte t?cnico final en la subcarpeta
    WordDoc.SaveAs2 carpetaSalida & "\SSIFO14-03 Informe t?cnico.docx"
    
    ' Guardar como PDF
    nombrePDF = carpetaSalida & "\SSIFO14-03 Informe t?cnico.pdf"
    WordDoc.ExportAsFixedFormat OutputFileName:= _
                                nombrePDF, ExportFormat:= _
                                17, OpenAfterExport:=True, OptimizeFor:= _
                                0, Range:=0, From:=1, To:=1, _
                                Item:=0, IncludeDocProps:=True, KeepIRM:=True, _
                                CreateBookmarks:=1, DocStructureTags:=True, _
                                BitmapMissingFonts:=True, UseISO19005_1:=False
    WordDoc.Close False
    
    ' Actualizar el documento de reporte ejecutivo
    Set WordDoc = WordApp.Documents.Open(tempDocPath2)
    
    Set rngReplace = WordDoc.content
    rngReplace.Find.ClearFormatting
    With rngReplace.Find
        .text = "«Total de vulnerabilidades»"
        .Replacement.text = totalVulnerabilidades
        .Forward = True
        .Wrap = 1        ' wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    
    If rngReplace.Find.Execute Then
        rngReplace.text = totalVulnerabilidades
    End If
    
    ' Actualizar el gr?fico InlineShape n?mero 1 en reporte ejecutivo
    FunActualizarGraficoSegunDicionario WordDoc, severityCounts, 1
    
    ' Actualizar el gr?fico InlineShape n?mero 2 en reporte ejecutivo
    FunActualizarGraficoSegunDicionario WordDoc, vulntypesCounts, 2
    
    ActualizarGraficos WordDoc
    
    ' Update the Table of Contents
    On Error Resume Next
    WordDoc.TablesOfContents(1).Update
    On Error GoTo 0
    
    ' Guardar el documento de reporte ejecutivo final en la subcarpeta
    WordDoc.SaveAs2 carpetaSalida & "\SSIFO15-03 Informe Ejecutivo.docx"
    ' Guardar como PDF
    nombrePDF = carpetaSalida & "\SSIFO14-03 Informe Ejecutivo.pdf"
    WordDoc.ExportAsFixedFormat OutputFileName:= _
                                nombrePDF, ExportFormat:= _
                                17, OpenAfterExport:=True, OptimizeFor:= _
                                0, Range:=0, From:=1, To:=1, _
                                Item:=0, IncludeDocProps:=True, KeepIRM:=True, _
                                CreateBookmarks:=1, DocStructureTags:=True, _
                                BitmapMissingFonts:=True, UseISO19005_1:=False
    WordDoc.Close False
    
    ' Cerrar la Aplicaci?n de Word
    WordApp.Quit
    Set WordDoc = Nothing
    Set WordApp = Nothing
    Set fileSystem = Nothing
    
    ' Mostrar mensaje de ?xito
    MsgBox "Se han generado los documentos de Word correctamente.", vbInformation
End Sub


Sub CYB003_GenerarReportesVulns()
    Dim selectedRange As Range
    Dim replaceDic  As Object
    Dim key         As Variant
    Dim value       As Variant
    Dim i           As Integer
    Dim headerRow   As Range
    Dim WordApp     As Object
    Dim WordDoc     As Object
    Dim plantillaReportePath As String
    Dim plantillaVulnerabilidadesPath As String
    Dim carpetaSalida As String
    Dim archivoTemp As String
    Dim fileSystem  As Object
    Dim dlg         As Object
    Dim tempFolder  As String
    Dim tempFolderGenerados As String
    Dim folderName  As String
    Dim secVulnerabilidades As String
    Dim rngReplace  As Object
    Dim tempFileName As String
    Dim documentsList() As String
    Dim numDocuments As Integer
    Dim finalDocumentPath As String
    Dim tempDocVulnerabilidadesPath As String


    ' Crear un diccionario para los reemplazos
    Set replaceDic = CreateObject("Scripting.Dictionary")
    
    ' Solicitar al usuario seleccionar el rango de celdas
    On Error Resume Next
    Set selectedRange = Application.InputBox("Seleccione una fila con los valores para procesar", Type:=8)
    On Error GoTo 0
    
    If selectedRange Is Nothing Then
        MsgBox "No se ha seleccionado un rango v?lido.", vbExclamation
        Exit Sub
    End If
    
    ' Verificar si el rango seleccionado pertenece a una tabla (ListObject)
    If Not selectedRange.ListObject Is Nothing Then
        ' Si es parte de una tabla, obtenemos el rango de la tabla
        Set tableRange = selectedRange.ListObject.Range
    Else
        ' Si no es parte de una tabla, mostrar un mensaje y salir
        MsgBox "El rango seleccionado no est? dentro de una tabla.", vbExclamation
        Exit Sub
    End If
    
    ' Buscar la fila de encabezados dentro del rango de la tabla
    For Each cell In tableRange.Rows(1).Cells
        If cell.value <> "" Then
            ' Encontramos la fila de encabezados
            Set headerRow = tableRange.Rows(1)
            Exit For
        End If
    Next cell
    
    If headerRow Is Nothing Then
        MsgBox "No se han encontrado encabezados en el rango seleccionado.", vbExclamation
        Exit Sub
    End If
    
    ' Obtener los encabezados y sus valores de la fila seleccionada
    For i = 1 To selectedRange.Columns.Count
        key = "«" & headerRow.Cells(1, i).value & "»"
        
        ' Verificar si la clave ya existe en el diccionario
        If replaceDic.Exists(key) Then
            MsgBox "Se ha encontrado un encabezado duplicado: " & headerRow.Cells(1, i).value & _
                   vbCrLf & "Por favor, corrige los encabezados duplicados y vuelve a ejecutar la macro.", vbExclamation
            Exit Sub
        End If
        
        ' Asignar el valor de la celda de la fila seleccionada al diccionario
        value = selectedRange.Cells(1, i).value
        replaceDic.Add key, value
    Next i
    
    ' Extraer el nombre de la Aplicaci?n
    If replaceDic.Exists("«Nombre de carpeta»") Then
        folderName = replaceDic("«Nombre de carpeta»")
    Else
        MsgBox "No se encontr? el campo        'Nombre de carpeta'.", vbExclamation
        Exit Sub
    End If
    
    ' Crear una subcarpeta con el nombre de la Aplicaci?n
    carpetaSalida = carpetaSalida & "\" & folderName
    On Error Resume Next
    MkDir carpetaSalida
    On Error GoTo 0
    
    If replaceDic.Exists("«Tipo de reporte»") Then
        Select Case replaceDic("«Tipo de reporte»")
            Case "T?cnico"
                
                ' Obtener la ruta de la plantilla directamente de la celda de la tabla
                If replaceDic.Exists("«Ruta de la plantilla»") Then
                    plantillaReportePath = replaceDic("«Ruta de la plantilla»")
                Else
                    MsgBox "No se encontr? el campo        'Ruta de la plantilla'.", vbExclamation
                    Exit Sub
                End If
                
                ' Verificar que la ruta de la plantilla exista
                If Len(Dir(plantillaReportePath)) = 0 Then
                    MsgBox "La ruta de la plantilla no es v?lida o el archivo no existe: " & plantillaReportePath, vbExclamation
                    Exit Sub
                End If
                
                Set dlg = Application.FileDialog(msoFileDialogFilePicker)
                ' Crear di?logos para seleccionar la carpeta de salida
                With Application.FileDialog(msoFileDialogFolderPicker)
                    .Title = "Seleccionar Carpeta de Salida"
                    If .Show = -1 Then
                        carpetaSalida = .SelectedItems(1)
                    Else
                        MsgBox "No se seleccion? ninguna carpeta. La macro se detendr?."
                        Exit Sub
                    End If
                End With
                
                Set dlg = Application.FileDialog(msoFileDialogFilePicker)
                dlg.Filters.Clear ' Borra los filtros existentes
                dlg.Filters.Add "Archivos de Word", "*.docx; *.doc; *.dotx; *.dot" ' Agrega un filtro para archivos de Word

                dlg.Title = "Seleccionar la plantilla de tabla de vulnerabilidades"
                If dlg.Show = -1 Then
                    plantillaVulnerabilidadesPath = dlg.SelectedItems(1)
                Else
                    MsgBox "No se seleccion? ningún archivo. La macro se detendr?."
                    Exit Sub
                End If
                
                ' Crear una instancia de Word
                On Error Resume Next
                Set WordApp = CreateObject("Word.Application")
                On Error GoTo 0
                If WordApp Is Nothing Then
                    MsgBox "No se puede iniciar Microsoft Word."
                    Exit Sub
                End If
                
                ' Crear una carpeta temporal
                tempFolder = Environ("TEMP") & "\tmp-" & Format(Now(), "yyyymmddhhmmss")
                On Error Resume Next
                MkDir tempFolder
                On Error GoTo 0
                
                ' Crear una carpeta para documentos generados
                tempFolderGenerados = tempFolder & "\Documentos_generados"
                On Error Resume Next
                MkDir tempFolderGenerados
                On Error GoTo 0
                
                ' Copiar la plantilla de reporte y procesar los documentos
                archivoTemp = tempFolder & "\" & Dir(plantillaReportePath)
                Set fileSystem = CreateObject("Scripting.FileSystemObject")
                fileSystem.CopyFile plantillaReportePath, archivoTemp
                
                Set WordDoc = WordApp.Documents.Open(archivoTemp)
                WordApp.Visible = False
                ReemplazarCampos WordDoc, replaceDic
                finalDocumentPath = carpetaSalida & "\" & "Informe T?cnico.docx"
                WordDoc.SaveAs finalDocumentPath
                WordDoc.Close
                Set WordDoc = Nothing
                WordApp.Quit
                Set WordApp = Nothing
                
            Case "Matriz de vulnerabilidades"
                
                ' Solicitar al usuario seleccionar el rango de celdas
                On Error Resume Next
                Set selectedRange = Application.InputBox("Seleccione el rango de celdas que contienen las columnas a considerar", Type:=8)
                On Error GoTo 0
                
                If selectedRange Is Nothing Then
                    MsgBox "No se ha seleccionado un rango v?lido.", vbExclamation
                    Exit Sub
                End If
                
                ' Verificar si el rango seleccionado pertenece a una tabla (ListObject)
                If Not selectedRange.ListObject Is Nothing Then
                    ' Si es parte de una tabla, obtenemos el rango de la tabla
                    Set tableRange = selectedRange.ListObject.Range
                Else
                    ' Si no es parte de una tabla, mostrar un mensaje y salir
                    MsgBox "El rango seleccionado no est? dentro de una tabla.", vbExclamation
                    Exit Sub
                End If
                
                Dim resultado As Boolean
                
                ' Llamar a la funci?n para exportar la hoja activa a Excel
                resultado = FunExportarHojaActivaAExcelINAI(carpetaSalida, folderName, tableRange.Worksheet, replaceDic("«Nombre del reporte»"))
                
            Case "Tablas de vulnerabilidades"
                
                GenerarDocumentosVulnerabilidiadesWord (replaceDic("«Nombre del reporte»"))
                
            Case Else
                MsgBox "El tipo de reporte no es reconocido.", vbExclamation
                Exit Sub
        End Select
    Else
        MsgBox "No se encontr? el campo        'Tipo de reporte'.", vbExclamation
        Exit Sub
    End If
    
    CYB039_KillAllWordInstances
    
    MsgBox "Proceso completado correctamente."
End Sub

Sub CYB039_KillAllWordInstances()
    Dim objWMI      As Object
    Dim objProcesses As Object
    Dim objProcess  As Object
    
    On Error Resume Next
    ' Get WMI service
    Set objWMI = GetObject("winmgmts:\\.\root\CIMV2")
    ' Retrieve all processes with name "WINWORD.EXE"
    Set objProcesses = objWMI.ExecQuery("SELECT * FROM Win32_Process WHERE Name =        'WINWORD.EXE'")
    
    ' Loop through each process and terminate it
    For Each objProcess In objProcesses
        objProcess.Terminate
    Next
    
    ' Clean up
    Set objProcess = Nothing
    Set objProcesses = Nothing
    Set objWMI = Nothing
    On Error GoTo 0
    
    MsgBox "All Microsoft Word instances have been terminated.", vbInformation
End Sub

Sub ReemplazarCampos(WordDoc As Object, replaceDic As Object)
    Dim key         As Variant
    Dim WordApp     As Object
    Dim docContent  As Object
    Dim findInRange As Boolean
    
    ' Obtener la aplicaci?n de Word
    Set WordApp = WordDoc.Application
    
    ' Obtener el contenido del documento
    Set docContent = WordDoc.content
    
    ' Bucle para buscar y reemplazar todas las ocurrencias en el diccionario
    For Each key In replaceDic.Keys
        ' Configurar la búsqueda
        With WordApp.Selection.Find
            .ClearFormatting
            .text = key
            .Forward = True
            .Wrap = 1        ' wdFindStop (detiene la búsqueda al final)
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            
            ' Intentar encontrar y reemplazar en todo el documento
            findInRange = .Execute
            Do While findInRange
                ' Realizar el reemplazo
                WordApp.Selection.text = CStr(replaceDic(key))
                ' Continuar buscando la siguiente ocurrencia
                findInRange = .Execute
            Loop
        End With
    Next key
    
    ' Limpiar objetos
    Set docContent = Nothing
    Set WordApp = Nothing
End Sub


Sub ActualizarGraficos(ByRef WordDoc As Object)
    ' Actualizar todos los gr?ficos en el documento de Word
    On Error Resume Next
    
    ' Recorrer todos los InlineShapes en el documento
    Dim i           As Integer
    Dim Chart       As Object
    Dim ChartData   As Object
    Dim ChartWorkbook As Object
    
    For i = 1 To WordDoc.InlineShapes.Count
        With WordDoc.InlineShapes(i)
            ' Verificar si el InlineShape es un gr?fico (wdInlineShapeChart = 12)
            If .Type = 12 And .HasChart Then
                Set Chart = .Chart
                If Not Chart Is Nothing Then
                    ' Activar los datos del gr?fico
                    Set ChartData = Chart.ChartData
                    If Not ChartData Is Nothing Then
                        ChartData.Activate
                        Set ChartWorkbook = ChartData.Workbook
                        If Not ChartWorkbook Is Nothing Then
                            ' Ocultar la ventana del libro de trabajo
                            ChartWorkbook.Application.Visible = False
                            ' Cerrar el libro de trabajo sin guardar cambios
                            ChartWorkbook.Close SaveChanges:=False
                        End If
                        ' Refrescar el gr?fico
                        Chart.Refresh
                    End If
                End If
            End If
        End With
    Next i
End Sub

Function GenerarDocumentosVulnerabilidiadesWord(fileName As String)
    Dim rng         As Range
    Dim tbl         As ListObject
    Dim WordApp     As Object
    Dim WordDoc     As Object
    Dim templatePath As String
    Dim outputPath  As String
    Dim replaceDic  As Object
    Dim cell        As Range
    Dim colIndex    As Integer
    Dim rowCount    As Integer
    Dim i           As Integer
    Dim tempFolder  As String
    Dim tempFolderPath As String
    Dim saveFolder  As String
    Dim selectedRange As Range        ' Variable para almacenar el rango seleccionado por el usuario
    Dim documentsList() As String        ' Lista para almacenar los documentos generados
    
    ' Solicita al usuario seleccionar el rango de celdas que contienen las columnas a considerar
    On Error Resume Next
    Set selectedRange = Application.InputBox("Seleccione el rango de celdas que contienen las columnas a considerar", Type:=8)
    On Error GoTo 0
    
    If selectedRange Is Nothing Then
        MsgBox "No se ha seleccionado un rango v?lido.", vbExclamation
        Exit Function
    End If
    
    ' Verifica si el rango seleccionado est? dentro de una tabla
    On Error Resume Next
    Set rng = selectedRange.ListObject.Range
    On Error GoTo 0
    
    If rng Is Nothing Then
        MsgBox "El rango seleccionado no est? dentro de una tabla.", vbExclamation
        Exit Function
    End If
    
    ' Solicita al usuario la ruta del documento de Word
    templatePath = Application.GetOpenFilename("Documentos de Word (*.docx), *.docx", , "Seleccione un documento de Word como plantilla")
    If templatePath = "Falso" Then Exit Function
    
    ' Solicita al usuario la carpeta donde desea guardar los archivos generados
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Seleccione la carpeta donde desea guardar los documentos generados"
        If .Show = -1 Then
            saveFolder = .SelectedItems(1)
        Else
            Exit Function
        End If
    End With
    
    ' Crea una instancia de la aplicaci?n de Word
    Set WordApp = CreateObject("Word.Application")
    WordApp.Visible = True
    
    ' Abre el documento de Word seleccionado
    Set WordDoc = WordApp.Documents.Open(templatePath)
    
    ' Crea un diccionario de reemplazo para los campos
    Set replaceDic = CreateObject("Scripting.Dictionary")
    
    ' Llena el diccionario de reemplazo con los datos de la tabla de Excel
    rowCount = rng.Rows.Count
    For Each cell In selectedRange.Rows(1).Cells        ' Tomamos la primera fila para los nombres de los campos
        replaceDic("«" & cell.value & "»") = ""
    Next cell
    
    ' Crea una carpeta temporal en la carpeta de archivos temporales del sistema
    tempFolder = Environ("TEMP") & "\tmp-" & Format(Now(), "yyyymmddhhmmss")
    MkDir tempFolder
    
    ' Copia el documento de Word seleccionado a la carpeta temporal
    Dim fs          As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    ' Genera un archivo de Word por cada registro de la tabla
    For i = 2 To rowCount        ' Empezamos desde la segunda fila para los datos reales
        ' Crear un nuevo diccionario para cada fila
        Set replaceDic = CreateObject("Scripting.Dictionary")
        
        ' Llena el diccionario de reemplazo con los datos de la fila actual de la tabla de Excel
        For Each cell In selectedRange.Rows(1).Cells        ' Tomamos la primera fila para los nombres de los campos
            replaceDic("«" & cell.value & "»") = rng.Cells(i, cell.Column).value
        Next cell
        
        ' Crea una copia del documento de Word en la carpeta temporal
        fs.CopyFile templatePath, tempFolder & "\Tabla_" & i & ".docx"
        ' Abre la copia del documento de Word
        Set WordDoc = WordApp.Documents.Open(tempFolder & "\Tabla_" & i & ".docx")
        ' Realiza los reemplazos en el documento de Word
        For Each key In replaceDic.Keys
            
            Debug.Print CStr(key)
            If CStr(key) = "«Descripci?n»" Then
                ' Aplicar la funci?n específica para la clave «Descripcion»
                replaceDic(key) = TransformarTexto(replaceDic(key))
            End If
            If CStr(key) = "«Propuesta de remediaci?n»" Then
                replaceDic(key) = TransformarTexto(replaceDic(key))
            End If
            If CStr(key) = "Referencias" Then
                replaceDic(key) = TransformarTexto(replaceDic(key))
            End If
            ' Reemplazar en el documento de Word
            WordAppReemplazarParrafo WordApp, WordDoc, CStr(key), CStr(replaceDic(key))
            
           
        Next key
        FormatearCeldaNivelRiesgo WordDoc.Tables(1).cell(1, 2)
        ' Guarda y cierra el documento de Word
        ' Antes de guardar el documento de Word
        'EliminarUltimasFilasSiEsSalidaPruebaSeguridad WordDoc, replaceDic
        WordDoc.Save
        WordDoc.Close
        
        ' Agregar el documento generado a la lista
        ReDim Preserve documentsList(i - 2)
        documentsList(i - 2) = tempFolder & "\Tabla_" & i & ".docx"
    Next i
    
    ' Combina todos los archivos en uno solo
    Dim finalDocumentPath As String
    finalDocumentPath = saveFolder & "\" & fileName & ".docx"
    FusionarDocumentosInsertando WordApp, documentsList, finalDocumentPath
    
    ' Mueve la carpeta temporal a la carpeta seleccionada por el usuario
    fs.MoveFolder tempFolder, saveFolder & "\Documentos_generados"
    
    ' Cerrar la aplicaci?n de Word
    WordApp.Quit
    Set WordApp = Nothing
    
    ' Muestra un mensaje de ?xito
    MsgBox "Se han generado los documentos de Word correctamente.", vbInformation
End Function



Sub FormatearParrafosGuionesCelda(cell As Object)
    Dim cellText As String
    Dim p As Object
    Dim rng As Object
    Dim posDosPuntos As Integer
    Dim strTexto As String

    ' Limpieza del texto en la celda
    cellText = Trim(Replace(cell.Range.text, vbCrLf, ""))
    cellText = Trim(Replace(cellText, vbCr, ""))
    cellText = Trim(Replace(cellText, vbLf, ""))
    cellText = Trim(Replace(cellText, Chr(7), ""))

    ' Recorre cada p?rrafo dentro de la celda
    For Each p In cell.Range.Paragraphs
        strTexto = p.Range.text
        
        ' Si el p?rrafo comienza con "- "
        If Left(Trim(strTexto), 2) = "- " Then
            posDosPuntos = InStr(strTexto, ":")
            
            ' Si hay dos puntos en el texto
            If posDosPuntos > 0 Then
                ' Aplicar negrita al texto antes de los dos puntos
                Set rng = p.Range
                rng.Start = p.Range.Start
                rng.End = p.Range.Start + posDosPuntos - 1
                rng.Font.Bold = True
                
                ' El texto despu?s de los dos puntos no tendr? negrita
                Set rng = p.Range
                rng.Start = p.Range.Start + posDosPuntos
                rng.End = p.Range.End
                rng.Font.Bold = False
            End If
        End If
    Next p
End Sub

Sub FormatearCeldaNivelRiesgo(cell As Object)
    Dim cellText As String
    Dim cvssScore As Double
    Dim isNumber As Boolean
    
    ' Obtener el texto de la celda y limpiar caracteres especiales
    cellText = Trim(Replace(cell.Range.text, vbCrLf, ""))
    cellText = Trim(Replace(cellText, vbCr, ""))
    cellText = Trim(Replace(cellText, vbLf, ""))
    cellText = Trim(Replace(cellText, Chr(7), ""))
    
    ' Intentar convertir el texto en un número
    On Error Resume Next
    cvssScore = CDbl(cellText)
    isNumber = (Err.Number = 0)
    On Error GoTo 0
    
    ' Si es un número, aplicar el formato según el rango del CVSS
    If isNumber Then
        Select Case cvssScore
            Case Is >= 9
                cell.Shading.BackgroundPatternColor = 10498160 ' CRÍTICA
                cell.Range.Font.Color = 16777215
            Case Is >= 7
                cell.Shading.BackgroundPatternColor = 255 ' ALTA
                cell.Range.Font.Color = 16777215
            Case Is >= 4
                cell.Shading.BackgroundPatternColor = 65535 ' MEDIA
                cell.Range.Font.Color = 0
            Case Is >= 0.1
                cell.Shading.BackgroundPatternColor = 5287936 ' BAJA
                cell.Range.Font.Color = 16777215
            Case Else
                cell.Shading.BackgroundPatternColor = wdColorAutomatic ' Sin color
        End Select
    Else
        ' Si no es un número, usar la clasificaci?n por texto
        Select Case UCase(cellText)
            Case "CRÍTICA"
                cell.Shading.BackgroundPatternColor = 10498160
                cell.Range.Font.Color = 16777215
            Case "ALTA"
                cell.Shading.BackgroundPatternColor = 255
                cell.Range.Font.Color = 16777215
            Case "MEDIA"
                cell.Shading.BackgroundPatternColor = 65535
                cell.Range.Font.Color = 0
            Case "BAJA"
                cell.Shading.BackgroundPatternColor = 5287936
                cell.Range.Font.Color = 16777215
        End Select
    End If
End Sub


Function TransformarTexto(text As String) As String
    Dim regex       As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    ' Configurar la expresi?n regular para encontrar saltos de línea o saltos de carro sin un punto antes y no seguidos de par?ntesis ni de gui?n
    With regex
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .pattern = "([^.()\r\n-])(?![^(]*\)|[-])[^\S\r\n]*[\r\n]+"        ' Expresi?n regular para encontrar saltos de línea o saltos de carro sin un punto antes y no seguidos de par?ntesis ni de gui?n
    End With
    
    ' Realizar la transformaci?n: quitar caracteres especiales y aplicar la expresi?n regular
    TransformarTexto = regex.Replace(Replace(text, Chr(7), ""), "$1 ")
End Function

Function FunActualizarGraficoSegunDicionario(ByRef WordDoc As Object, conteos As Object, graficoIndex As Integer) As Boolean
    Dim ils         As Object
    Dim Chart       As Object
    Dim ChartData   As Object
    Dim ChartWorkbook As Object
    Dim SourceSheet As Object
    Dim dataRangeAddress As String
    Dim categoryRow As Integer
    Dim category    As Variant
    Dim lastRow     As Long
    Dim tableIndex  As Integer
    Dim sheetIndex  As Integer
    
    tableIndex = 1
    sheetIndex = 1
    
    On Error GoTo ErrorHandler
    
    ' Verificar que el ?ndice del gr?fico es v?lido
    If graficoIndex < 1 Or graficoIndex > WordDoc.InlineShapes.Count Then
        MsgBox "Índice de gr?fico fuera de rango."
        FunActualizarGraficoSegunDicionario = False
        Exit Function
    End If
    
    ' Obtener el InlineShape correspondiente al ?ndice
    Set ils = WordDoc.InlineShapes(graficoIndex)
    
    If ils.Type = 12 And ils.HasChart Then
        Set Chart = ils.Chart
        If Not Chart Is Nothing Then
            ' Activar el libro de trabajo asociado al gr?fico
            Set ChartData = Chart.ChartData
            If Not ChartData Is Nothing Then
                ChartData.Activate
                Set ChartWorkbook = ChartData.Workbook
                If Not ChartWorkbook Is Nothing Then
                    Set SourceSheet = ChartWorkbook.Sheets(1)        ' Usar la primera hoja del libro
                    
                    ' Limpiar las celdas en el rango antes de agregar nuevos datos
                    lastRow = SourceSheet.Cells(SourceSheet.Rows.Count, 1).End(xlUp).Row
                    If lastRow >= 2 Then
                        SourceSheet.Range("A2:B" & lastRow).ClearContents
                    End If
                    
                    ' Insertar nuevos datos
                    categoryRow = 2        ' Empezar en la fila 2 para los tipos de vulnerabilidad
                    For Each category In conteos.Keys
                        SourceSheet.Cells(categoryRow, 1).value = category
                        SourceSheet.Cells(categoryRow, 2).value = conteos(category)
                        categoryRow = categoryRow + 1
                    Next category
                    
                    ' Construir el rango din?mico como una cadena
                    dataRangeAddress =        '" & SourceSheet.Name & "'!$A$1:$B$" & (categoryRow - 1)
                    Debug.Print dataRangeAddress
                    
                    ' Verifica si la tabla existe usando el ?ndice
                    On Error Resume Next
                    Set DataTable = SourceSheet.ListObjects(tableIndex)        ' Obtiene el objeto de la tabla por ?ndice
                    On Error GoTo 0
                    
                    ' Verifica que el objeto de la tabla no sea Nothing
                    If Not DataTable Is Nothing Then
                        ' Redimensiona la tabla al nuevo rango usando el objeto Worksheet
                        DataTable.Resize SourceSheet.Range("A1:B" & (categoryRow - 1))
                    Else
                        MsgBox "La tabla en el índice " & tableIndex & " no se encontr? en la hoja."
                    End If
                    
                    WordDoc.InlineShapes(graficoIndex).Chart.SetSourceData Source:=dataRangeAddress
                    
                    ' Actualizar el gr?fico
                    Chart.Refresh
                    
                    ' Cerrar el libro de trabajo sin guardar cambios
                    ChartWorkbook.Close SaveChanges:=False
                    
                    FunActualizarGraficoSegunDicionario = True
                End If
            End If
        Else
            MsgBox "El InlineShape seleccionado no contiene un gr?fico v?lido."
            FunActualizarGraficoSegunDicionario = False
        End If
    Else
        MsgBox "El InlineShape seleccionado no contiene un gr?fico."
        FunActualizarGraficoSegunDicionario = False
    End If
    
    Exit Function
    
ErrorHandler:
    MsgBox "Ocurri? un error: " & Err.Description, vbCritical
    FunActualizarGraficoSegunDicionario = False
End Function




Sub WordAppReemplazarParrafo(WordApp As Object, WordDoc As Object, wordToFind As String, replaceWord As String)
    Dim findInRange As Boolean
    Dim searchRange As Object
    
    ' Configurar el rango para todo el documento
    Set searchRange = WordDoc.content
    
    ' Configurar las opciones de búsqueda
    With searchRange.Find
        .text = wordToFind
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    
    ' Bucle para buscar y reemplazar todas las ocurrencias
    Do While searchRange.Find.Execute
        ' Reemplazar el texto encontrado con la cadena larga
        searchRange.text = replaceWord
        
        ' Mover el rango al siguiente texto para evitar bucles infinitos
        searchRange.Collapse Direction:=wdCollapseEnd
    Loop
End Sub



Sub CYB013_DesglosarIPs()
    Dim celda As Range
    Dim ipRango As String
    Dim partes() As String
    Dim ipInicio As String, ipFin As String
    Dim numInicio As Integer, numFin As Integer
    Dim i As Integer
    Dim filaActual As Integer

    ' Obtener la celda seleccionada
    Set celda = Selection

    ' Verificar si la celda no est? vacía
    If IsEmpty(celda) Then
        MsgBox "Seleccione una celda con un rango de IPs.", vbExclamation, "Error"
        Exit Sub
    End If

    ipRango = Trim(celda.value)

    ' Dividir el rango de IPs usando el guion como separador
    partes = Split(ipRango, "-")
    
    ' Verificar que haya dos partes en el rango
    If UBound(partes) <> 1 Then
        MsgBox "Formato inv?lido. Use: 10.0.1.60-10.0.1.78", vbExclamation, "Error"
        Exit Sub
    End If

    ipInicio = partes(0)
    ipFin = partes(1)

    ' Extraer el último número de las IPs
    numInicio = CInt(Split(ipInicio, ".")(3))
    numFin = CInt(Split(ipFin, ".")(3))

    ' Validar que el inicio es menor o igual que el fin
    If numInicio > numFin Then
        MsgBox "El rango de IPs es inv?lido.", vbExclamation, "Error"
        Exit Sub
    End If

    ' Obtener la parte fija de la IP (sin el último octeto)
    Dim baseIP As String
    baseIP = Left(ipInicio, InStrRev(ipInicio, "."))

    ' Insertar las IPs debajo de la celda seleccionada
    filaActual = celda.Row + 1

    For i = numInicio To numFin
        Cells(filaActual, celda.Column).value = baseIP & i
        filaActual = filaActual + 1
    Next i

    MsgBox "Rango de IPs desglosado correctamente.", vbInformation, "Completado"
End Sub




Sub CYB034_CargarResultados_DatosDesdeCSVNessus()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim archivos As Variant
    Dim i As Integer
    Dim wbCSV As Workbook
    Dim wsCSV As Worksheet
    Dim encabezados As Variant
    Dim csvData As Variant
    Dim fila As ListRow
    Dim colCSVIndex As Integer
    Dim columnaDestino As Integer
    Dim columnaCorrespondiente As String
    
    ' Asignar la hoja de trabajo activa y la tabla seleccionada
    Set wb = ThisWorkbook
    Set ws = wb.ActiveSheet
    
    ' Verificar si la celda activa est? dentro de una tabla
    Dim celdaEnTabla As Boolean
    celdaEnTabla = False
    For Each tbl In ws.ListObjects
        If Not tbl.DataBodyRange Is Nothing Then
            If Not Intersect(ActiveCell, tbl.DataBodyRange) Is Nothing Or Not Intersect(ActiveCell, tbl.Range) Is Nothing Then
                celdaEnTabla = True
                Exit For
            End If
        End If
    Next tbl
    
    ' Si no est? dentro de una tabla, salir
    If Not celdaEnTabla Then
        MsgBox "La celda seleccionada no est? dentro de una tabla.", vbExclamation
        Exit Sub
    End If
    
    ' Seleccionar múltiples archivos CSV
    archivos = Application.GetOpenFilename("Archivos CSV (*.csv), *.csv", MultiSelect:=True, Title:="Seleccionar archivos CSV")
    If IsArray(archivos) = False Then Exit Sub
    
    ' Procesar cada archivo seleccionado
    For i = LBound(archivos) To UBound(archivos)
        ' Abrir el archivo CSV
        Set wbCSV = Workbooks.Open(fileName:=archivos(i), Local:=True)
        Set wsCSV = wbCSV.Sheets(1)
        
        ' Leer los encabezados y datos
        encabezados = wsCSV.UsedRange.Rows(1).value
        csvData = wsCSV.UsedRange.Offset(1, 0).value
        
        ' Cerrar el archivo CSV
        wbCSV.Close False
        
        ' Cargar los datos en la tabla de Excel
        Dim j As Integer
        For j = 1 To UBound(csvData, 1)
            Set fila = tbl.ListRows.Add
            
            For colCSVIndex = 1 To UBound(encabezados, 2)
                ' Mapeo de columnas
                Select Case encabezados(1, colCSVIndex)
                    Case "Host": columnaCorrespondiente = "IPv4 Interna"
                    Case "CVE": columnaCorrespondiente = "CVE"
                    Case "CVSS v3.0 Base Score": columnaCorrespondiente = "CVSSScore"
                    Case "Description": columnaCorrespondiente = "Descripci?n ampliada"
                    Case "Metasploit": columnaCorrespondiente = "Exploits públicos"
                    Case "Plugin ID": columnaCorrespondiente = "Identificador original de la vulnerabilidad"
                    Case "Name": columnaCorrespondiente = "Nombre de vulnerabilidad"
                    Case "Protocol": columnaCorrespondiente = "Protocolo de transporte"
                    Case "Port": columnaCorrespondiente = "Puerto"
                    Case "See Also": columnaCorrespondiente = "Referencias"
                    Case "Plugin Output": columnaCorrespondiente = "Salidas de herramienta"
                    Case "Risk": columnaCorrespondiente = "Severidad"
                    Case Else: columnaCorrespondiente = ""
                End Select
                
                ' Si la columna existe en la tabla, asignar el valor
                If columnaCorrespondiente <> "" Then
                    columnaDestino = tbl.ListColumns(columnaCorrespondiente).Index
                    fila.Range(1, columnaDestino).value = csvData(j, colCSVIndex)
                End If
            Next colCSVIndex
            
            ' Colocar el tipo de origen
            columnaDestino = tbl.ListColumns("Tipo de origen").Index
            fila.Range(1, columnaDestino).value = "Nessus"
        Next j
    Next i
    
    MsgBox "Datos cargados con ?xito en la tabla.", vbInformation
End Sub


Sub CYB035_CargarResultados_DatosDesdeCSVNexPose()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim archivos As Variant
    Dim archivo As Variant
    Dim mensaje As String
    Dim respuesta As VbMsgBoxResult
    Dim wbCSV As Workbook
    Dim wsCSV As Worksheet
    Dim csvData As Variant
    Dim encabezados As Variant
    Dim i As Integer, colCSVIndex As Integer, columnaDestino As Integer
    Dim fila As ListRow
    Dim columnaCorrespondiente As String
    
    ' Asignar la hoja de trabajo activa y la tabla seleccionada
    Set wb = ThisWorkbook
    Set ws = wb.ActiveSheet
    
    ' Verificar si la celda activa est? dentro de una tabla
    Dim celdaEnTabla As Boolean
    celdaEnTabla = False
    For Each tbl In ws.ListObjects
        If Not tbl.DataBodyRange Is Nothing Then
            If Not Intersect(ActiveCell, tbl.DataBodyRange) Is Nothing Or Not Intersect(ActiveCell, tbl.Range) Is Nothing Then
                celdaEnTabla = True
                Exit For
            End If
        End If
    Next tbl
    
    ' Si no est? dentro de una tabla, salir
    If Not celdaEnTabla Then
        MsgBox "La celda seleccionada no est? dentro de una tabla.", vbExclamation
        Exit Sub
    End If
    
    ' Preguntar al usuario si est? seguro de cargar los datos
    mensaje = "¿Est? seguro que desea cargar datos de los archivos CSV en la tabla '" & tbl.Name & "'?"
    respuesta = MsgBox(mensaje, vbYesNo + vbQuestion, "Confirmaci?n")
    If respuesta = vbNo Then Exit Sub
    
    ' Seleccionar múltiples archivos CSV
    archivos = Application.GetOpenFilename("Archivos CSV (*.csv), *.csv", MultiSelect:=True, Title:="Seleccionar archivos CSV")
    If Not IsArray(archivos) Then Exit Sub ' Si el usuario cancela la selecci?n
    
    ' Iterar sobre los archivos seleccionados
    For Each archivo In archivos
        ' Verificar que el archivo seleccionado es CSV
        If LCase(Trim(Right(archivo, 4))) <> ".csv" Then
            MsgBox "El archivo " & archivo & " no es un archivo CSV.", vbExclamation
            Exit Sub
        End If
        
        ' Abrir el archivo CSV
        Set wbCSV = Workbooks.Open(fileName:=archivo, Local:=True)
        Set wsCSV = wbCSV.Sheets(1) ' Asumimos que los datos est?n en la primera hoja
        
        ' Leer los encabezados desde la primera fila del archivo CSV
        encabezados = wsCSV.UsedRange.Rows(1).value
        
        ' Leer los datos del CSV
        csvData = wsCSV.UsedRange.Offset(1, 0).value
        
        ' Cerrar el archivo CSV sin guardar cambios
        wbCSV.Close False
        
        ' Cargar los datos en la tabla de Excel
        For i = 1 To UBound(csvData, 1) ' Recorrer filas del CSV
            Set fila = tbl.ListRows.Add ' Agregar nueva fila a la tabla
            
            For colCSVIndex = 1 To UBound(encabezados, 2) ' Recorrer columnas del CSV
                Select Case encabezados(1, colCSVIndex)
                    Case "Asset IP Address"
                        tbl.ListColumns("IPv4 Interna").DataBodyRange.Cells(tbl.ListRows.Count, 1).value = csvData(i, colCSVIndex)
                        tbl.ListColumns("Identificador de detecci?n usado").DataBodyRange.Cells(tbl.ListRows.Count, 1).value = csvData(i, colCSVIndex)
                    Case "Service Port"
                        columnaCorrespondiente = "Puerto"
                    Case "Vulnerability Title"
                        columnaCorrespondiente = "Nombre de vulnerabilidad"
                        tbl.ListColumns("Identificador original de la vulnerabilidad").DataBodyRange.Cells(tbl.ListRows.Count, 1).value = csvData(i, colCSVIndex)
                    Case "Vulnerability Severity Level"
                        columnaCorrespondiente = "Severidad"
                    Case Else
                        columnaCorrespondiente = ""
                End Select
                
                ' Si hay una columna correspondiente, asignar el valor
                If columnaCorrespondiente <> "" Then
                    columnaDestino = tbl.ListColumns(columnaCorrespondiente).Index
                    fila.Range(1, columnaDestino).value = csvData(i, colCSVIndex)
                End If
            Next colCSVIndex
            
            ' Asignar valor fijo a "Tipo de origen"
            fila.Range(1, tbl.ListColumns("Tipo de origen").Index).value = "Nexpose"
        Next i
    Next archivo
    
    MsgBox "Datos cargados con ?xito en la tabla.", vbInformation
End Sub


Sub CYB036_CargarResultados_DatosDesdeXMLOpenVAS()
    Dim wb As Workbook, ws As Worksheet, tbl As ListObject
    Dim archivo As String
    Dim xmlDoc As Object, resultNodes As Object, resultNode As Object
    Dim respuesta As Integer, mensaje As String
    Dim fila As ListRow
    Dim dict As Object, header As Range
    Dim requiredFields As Variant, field As Variant
Dim regex As Object
Set regex = CreateObject("VBScript.RegExp")

    ' Asignar la hoja de trabajo activa
    Set wb = ThisWorkbook
    Set ws = wb.ActiveSheet
    
    ' Verificar si la celda activa est? dentro de alguna tabla y asignarla a tbl
    Dim celdaEnTabla As Boolean, t As ListObject
    celdaEnTabla = False
    For Each t In ws.ListObjects
        If Not t.DataBodyRange Is Nothing Then
            If Not Intersect(ActiveCell, t.Range) Is Nothing Then
                Set tbl = t
                celdaEnTabla = True
                Exit For
            End If
        End If
    Next t
    If Not celdaEnTabla Then
        MsgBox "La celda seleccionada no est? dentro de una tabla.", vbExclamation
        Exit Sub
    End If
    
    ' Confirmar con el usuario
    mensaje = "¿Est? seguro que desea cargar datos del archivo XML en la tabla '" & tbl.Name & "'?"
    respuesta = MsgBox(mensaje, vbYesNo + vbQuestion, "Confirmaci?n")
    If respuesta = vbNo Then Exit Sub
    
    ' Seleccionar el archivo XML
    archivo = Application.GetOpenFilename("Archivos XML (*.xml), *.xml", , "Seleccionar archivo XML")
    If archivo = "False" Then Exit Sub
    If LCase(Trim(Right(archivo, 4))) <> ".xml" Then
        MsgBox "El archivo seleccionado no es un archivo XML.", vbExclamation
        Exit Sub
    End If
    
    ' Cargar el archivo XML
    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    xmlDoc.async = False
    xmlDoc.Load archivo
    If xmlDoc.parseError.ErrorCode <> 0 Then
        MsgBox "Error al cargar el archivo XML: " & xmlDoc.parseError.Reason, vbExclamation
        Exit Sub
    End If
    
    ' Seleccionar los nodos <result> en la ruta /report/report/results/result
    Set resultNodes = xmlDoc.SelectNodes("/report/report/results/result")
    If resultNodes.Length = 0 Then
        MsgBox "No se encontraron registros en el XML.", vbExclamation
        Exit Sub
    End If
    
    ' Crear un diccionario para mapear los encabezados de la tabla a sus índices relativos
    Set dict = CreateObject("Scripting.Dictionary")
    For Each header In tbl.HeaderRowRange.Cells
        dict(Trim(header.value)) = header.Column - tbl.Range.Cells(1, 1).Column + 1
    Next header
    
    ' Definir los campos requeridos en la tabla
    requiredFields = Array("Severidad", "Nombre de vulnerabilidad", "Salidas de herramienta", "IPv4 Interna", "Puerto")
    For Each field In requiredFields
        If Not dict.Exists(field) Then
            MsgBox "La columna '" & field & "' no se encontr? en la tabla.", vbExclamation
            Exit Sub
        End If
    Next field
    
 ' Configure regex to keep only numbers
With regex
    .pattern = "\D"  ' Match any non-digit character
    .Global = True   ' Apply globally to the string
End With

' Recorrer cada nodo <result> y agregar una nueva fila a la tabla con los datos
For Each resultNode In resultNodes
    Set fila = tbl.ListRows.Add
    On Error Resume Next
    fila.Range.Cells(1, dict("Severidad")).value = Trim(resultNode.SelectSingleNode("severity").text)
    fila.Range.Cells(1, dict("Nombre de vulnerabilidad")).value = Trim(resultNode.SelectSingleNode("name").text)
    fila.Range.Cells(1, dict("Salidas de herramienta")).value = Trim(resultNode.SelectSingleNode("description").text)
    fila.Range.Cells(1, dict("IPv4 Interna")).value = Trim(resultNode.SelectSingleNode("host").text)

    ' Extract port number only
    fila.Range.Cells(1, dict("Puerto")).value = regex.Replace(Trim(resultNode.SelectSingleNode("port").text), "")

    On Error GoTo 0
Next resultNode

    MsgBox "Datos cargados con ?xito.", vbInformation
End Sub


Sub CYB037_CargarResultados_DatosDesdeCSVAcunetix()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim archivos As Variant
    Dim i As Integer
    Dim wbCSV As Workbook
    Dim wsCSV As Worksheet
    Dim encabezados As Variant
    Dim csvData As Variant
    Dim fila As ListRow              ' Para la fila NUEVA a agregar
    Dim filaExistente As ListRow ' Para la fila existente si se encuentra duplicado
    Dim lrExisting As ListRow      ' Para iterar en la búsqueda
    Dim colCSVIndex As Integer
    ' Dim columnaDestino As Integer ' Variable no usada, se puede eliminar
    Dim j As Long                  ' Usar Long para números de fila grandes
    Dim celdaEnTabla As Boolean
    Dim t As ListObject
    Dim mensaje As String
    Dim respuesta As VbMsgBoxResult

    ' Variables para almacenar valores específicos de la fila CSV
    Dim csvTarget As String
    Dim csvAffects As String
    Dim csvName As String
    Dim identificadorDeteccion As String
    Dim affectsProcessed As String

    ' Variables para la lógica de duplicados
    Dim registroEncontrado As Boolean
    Dim fechaActual As String
    Dim colIdxIdDeteccion As Long, colIdxTipoOrigen As Long, colIdxIdVuln As Long, colIdxNomVuln As Long
    Dim colIdxConteo As Long, colIdxFecha As Long
    Dim match1 As Boolean, match2 As Boolean, match3 As Boolean, match4 As Boolean
    Dim conteoActual As Variant
    Dim nuevoConteo As Long

    ' Variables para la validación específica de columnas
    Dim columnasFaltantes As String
    Dim colNombre As String

    ' Manejo básico de errores general
    On Error GoTo ErrorHandler

    ' Obtener fecha actual una vez (o por archivo si se prefiere)
    fechaActual = Format(Date, "dd/mm/yyyy") ' Asegura formato correcto

    ' Asignar la hoja de trabajo activa
    Set wb = ThisWorkbook
    Set ws = wb.ActiveSheet

    ' Verificar si la celda activa está dentro de una tabla y asignarla a tbl
    celdaEnTabla = False
    For Each t In ws.ListObjects
        If Not t.DataBodyRange Is Nothing Then
            ' Check if the table is actually visible and has cells
            On Error Resume Next ' Handle potential errors if table range is invalid
            Dim tblRangeAddress As String
            tblRangeAddress = t.Range.Address
            If Err.Number = 0 Then
                 If Not Intersect(ActiveCell, t.Range) Is Nothing Then
                    Set tbl = t
                    celdaEnTabla = True
                    Exit For
                 End If
            Else
                Debug.Print "Advertencia: Error al acceder al rango de la tabla '" & t.Name & "'. Error: " & Err.Description
                Err.Clear
            End If
            On Error GoTo ErrorHandler ' Restore main error handler
        End If
    Next t


    ' Si no está dentro de una tabla, salir
    If Not celdaEnTabla Then
        MsgBox "La celda seleccionada no está dentro de una tabla válida.", vbExclamation
        Exit Sub
    End If

    ' --- Inicio: Validación específica de columnas CLAVE ---
    columnasFaltantes = ""
    ' Usar una función auxiliar para no repetir código On Error


    ' Activar el manejador de errores principal por si algo falla ANTES de las comprobaciones
    On Error GoTo ErrorHandler

    ' Comprobar cada columna requerida individualmente
    CheckColumnExists tbl, "Identificador de detección usado", columnasFaltantes, colIdxIdDeteccion
    CheckColumnExists tbl, "Tipo de origen", columnasFaltantes, colIdxTipoOrigen
    CheckColumnExists tbl, "Identificador original de la vulnerabilidad", columnasFaltantes, colIdxIdVuln
    CheckColumnExists tbl, "Nombre original de la vulnerabilidad", columnasFaltantes, colIdxNomVuln
    ' *** CORREGIDO: Usar nombres exactos de tu tabla Excel ***
    CheckColumnExists tbl, "Conteo de detección", columnasFaltantes, colIdxConteo
    CheckColumnExists tbl, "Última fecha de detección", columnasFaltantes, colIdxFecha

    ' Si alguna columna faltó, mostrar el mensaje específico y salir
    If Len(columnasFaltantes) > 0 Then
        MsgBox "Error Crítico: La(s) siguiente(s) columna(s) requerida(s) no existe(n) en la tabla '" & tbl.Name & "':" & vbCrLf & vbCrLf & _
               columnasFaltantes & vbCrLf & vbCrLf & _
               "Por favor, asegúrese de que estas columnas existan con el nombre EXACTO (incluyendo tildes, mayúsculas/minúsculas si aplica y sin espacios extra).", _
               vbCritical, "Columnas Faltantes Específicas"
        Exit Sub ' Detener la ejecución
    End If
    ' --- Fin: Validación específica de columnas CLAVE ---

    ' Si todas las columnas existen, continuar con la confirmación y el proceso...
    ' Restaurar manejo de errores general por si acaso
    On Error GoTo ErrorHandler

    ' Confirmar con el usuario
    mensaje = "Esta macro cargará datos de Acunetix en '" & tbl.Name & "'." & vbCrLf & _
              "Verificará duplicados basados en las 4 columnas clave." & vbCrLf & _
              "- Si es nuevo: Agrega registro, Conteo=1, Fecha=Hoy." & vbCrLf & _
              "- Si existe: Incrementa 'Conteo de detección', actualiza 'Última fecha de detección'." & vbCrLf & vbCrLf & _
              "¿Desea continuar?"
    respuesta = MsgBox(mensaje, vbYesNo + vbQuestion, "Confirmación de Carga con Verificación")
    If respuesta = vbNo Then Exit Sub

    ' Seleccionar múltiples archivos CSV
    archivos = Application.GetOpenFilename("Archivos CSV (*.csv), *.csv", MultiSelect:=True, Title:="Seleccionar archivos CSV de Acunetix")
    If IsArray(archivos) = False Then Exit Sub ' Usuario canceló

    ' Desactivar actualizaciones de pantalla para mejorar rendimiento
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual ' Optimización adicional

    ' Procesar cada archivo seleccionado
    For i = LBound(archivos) To UBound(archivos)
        ' Abrir el archivo CSV
        ' Añadir manejo de error específico para la apertura del CSV
        On Error Resume Next ' Temporalmente para detectar si el archivo no se puede abrir
        Set wbCSV = Workbooks.Open(fileName:=archivos(i), Local:=True, ReadOnly:=True, Format:=6, Delimiter:=",") ' Format 6 = CSV, Delimiter puede necesitar ajuste
        If Err.Number <> 0 Then
            MsgBox "Error al abrir el archivo CSV: '" & Mid(archivos(i), InStrRev(archivos(i), "\") + 1) & "'." & vbCrLf & "Detalle: " & Err.Description & vbCrLf & "Saltando este archivo.", vbWarning
            Err.Clear
            On Error GoTo ErrorHandler ' Restaurar manejador principal
            GoTo SiguienteArchivo ' Saltar al siguiente archivo
        End If
        On Error GoTo ErrorHandler ' Restaurar manejador principal

        Set wsCSV = wbCSV.Sheets(1)

        ' Leer los encabezados y datos
        If wsCSV.FilterMode Then wsCSV.ShowAllData
        If Not wsCSV.UsedRange Is Nothing Then
             If wsCSV.UsedRange.Rows.Count > 0 Then
                 encabezados = wsCSV.UsedRange.Rows(1).value
             Else
                 MsgBox "El archivo CSV '" & Mid(archivos(i), InStrRev(archivos(i), "\") + 1) & "' está vacío o no tiene encabezados.", vbInformation
                 GoTo CerrarYSaltar ' Usar etiqueta para asegurar cierre
             End If
             If wsCSV.UsedRange.Rows.Count > 1 Then
                  ' Cuidado con tablas muy grandes, esto carga todo a memoria
                  csvData = wsCSV.UsedRange.Offset(1, 0).Resize(wsCSV.UsedRange.Rows.Count - 1, wsCSV.UsedRange.Columns.Count).value
             Else
                  csvData = Null ' Marcar como que no hay datos
             End If
        Else
             MsgBox "El archivo CSV '" & Mid(archivos(i), InStrRev(archivos(i), "\") + 1) & "' parece estar completamente vacío.", vbInformation
             GoTo CerrarYSaltar ' Usar etiqueta para asegurar cierre
        End If

        ' Cerrar el archivo CSV ANTES de procesar los datos en memoria
        wbCSV.Close False
        Set wsCSV = Nothing
        Set wbCSV = Nothing

        If IsNull(csvData) Then GoTo SiguienteArchivo ' Saltar si no había datos (ya está cerrado)

        ' --- Procesar filas del CSV ---
        For j = 1 To UBound(csvData, 1)
            ' Reiniciar variables para la fila CSV actual
            csvTarget = ""
            csvAffects = ""
            csvName = ""
            identificadorDeteccion = ""
            affectsProcessed = ""

            ' 1. Extraer valores necesarios del CSV (añadir manejo de error si columna CSV falta)
            Dim colTargetIdx As Integer: colTargetIdx = 0
            Dim colAffectsIdx As Integer: colAffectsIdx = 0
            Dim colNameIdx As Integer: colNameIdx = 0
            On Error Resume Next ' Buscar índices de columnas CSV
            For colCSVIndex = 1 To UBound(encabezados, 2)
                Select Case Trim(CStr(encabezados(1, colCSVIndex))) ' Convertir a String por si acaso
                    Case "Target": colTargetIdx = colCSVIndex
                    Case "Affects": colAffectsIdx = colCSVIndex
                    Case "Name": colNameIdx = colCSVIndex
                End Select
            Next colCSVIndex
            Err.Clear
            On Error GoTo ErrorHandler

            ' Validar si se encontraron las columnas CSV necesarias
            If colTargetIdx = 0 Or colAffectsIdx = 0 Or colNameIdx = 0 Then
                 Debug.Print "Advertencia: Faltan columnas ('Target', 'Affects', 'Name') en CSV: " & Mid(archivos(i), InStrRev(archivos(i), "\") + 1) & ". Saltando fila " & j
                 GoTo SiguienteFilaCSV ' Saltar esta fila si faltan datos clave del CSV
            End If

            ' Extraer datos usando los índices encontrados (con CStr para evitar errores de tipo)
            On Error Resume Next ' Para manejar celdas vacías o con error en CSV
            csvTarget = CStr(csvData(j, colTargetIdx))
            csvAffects = CStr(csvData(j, colAffectsIdx))
            csvName = CStr(csvData(j, colNameIdx))
            If Err.Number <> 0 Then
                Debug.Print "Advertencia: Error leyendo datos de fila " & j & " en CSV: " & Mid(archivos(i), InStrRev(archivos(i), "\") + 1) & ". Usando valores vacíos."
                csvTarget = "": csvAffects = "": csvName = ""
                Err.Clear
            End If
            On Error GoTo ErrorHandler


            ' 2. Calcular 'Identificador de detección usado' con la regla condicional
            If Right(csvTarget, 1) = "/" Then
                If Left(csvAffects, 1) = "/" Then
                    If Len(csvAffects) > 1 Then affectsProcessed = Mid(csvAffects, 2) Else affectsProcessed = ""
                Else
                    affectsProcessed = csvAffects
                End If
            Else
                 ' Si Target NO termina en "/", Affects se usa tal cual
                 affectsProcessed = csvAffects
            End If
             ' Concatenar SIEMPRE Target con el affects procesado
            identificadorDeteccion = csvTarget & affectsProcessed


            ' 3. Buscar si el registro ya existe en la tabla Excel
            registroEncontrado = False
            Set filaExistente = Nothing
            If tbl.ListRows.Count > 0 Then
                For Each lrExisting In tbl.ListRows
                    match1 = False: match2 = False: match3 = False: match4 = False
                    On Error Resume Next ' Ignorar error si una celda está vacía/error durante la comparación
                    ' Comparar los 4 valores clave usando CStr para manejar distintos tipos de datos
                    match1 = (CStr(lrExisting.Range(1, colIdxIdDeteccion).value) = identificadorDeteccion)
                    match2 = (CStr(lrExisting.Range(1, colIdxTipoOrigen).value) = "Acunetix")
                    match3 = (CStr(lrExisting.Range(1, colIdxIdVuln).value) = csvName)
                    match4 = (CStr(lrExisting.Range(1, colIdxNomVuln).value) = csvName)
                    ' Limpiar error potencial de CStr y restaurar manejo normal
                    If Err.Number <> 0 Then Err.Clear
                    On Error GoTo ErrorHandler

                    If match1 And match2 And match3 And match4 Then
                        registroEncontrado = True
                        Set filaExistente = lrExisting
                        Exit For
                    End If
                Next lrExisting
            End If

            ' 4. Actuar según si se encontró o no el registro
            If registroEncontrado Then
                ' --- REGISTRO EXISTENTE ---
                ' Incrementar Conteo de detección
                On Error Resume Next ' Manejar posible error al leer/escribir conteo
                conteoActual = filaExistente.Range(1, colIdxConteo).value
                If IsNumeric(conteoActual) And Not IsEmpty(conteoActual) And Not IsNull(conteoActual) Then
                    nuevoConteo = CLng(conteoActual) + 1
                Else
                    nuevoConteo = 1 ' Si está vacío, es error o no numérico, empezar en 1
                End If
                filaExistente.Range(1, colIdxConteo).value = nuevoConteo
                If Err.Number <> 0 Then
                    Debug.Print "Advertencia: No se pudo actualizar 'Conteo de detección' para fila existente (ID Detección: " & identificadorDeteccion & "). Error: " & Err.Description
                    Err.Clear
                End If
                On Error GoTo ErrorHandler ' Restaurar manejo normal

                ' Actualizar Última fecha de detección
                On Error Resume Next
                filaExistente.Range(1, colIdxFecha).value = CDate(fechaActual) ' Intentar convertir a Fecha para formato correcto
                 If Err.Number <> 0 Then
                    Err.Clear
                    filaExistente.Range(1, colIdxFecha).value = fechaActual ' Si falla CDate, usar texto
                    If Err.Number <> 0 Then
                       Debug.Print "Advertencia: No se pudo actualizar 'Última fecha de detección' para fila existente (ID Detección: " & identificadorDeteccion & "). Error: " & Err.Description
                       Err.Clear
                    End If
                 End If
                On Error GoTo ErrorHandler ' Restaurar manejo normal

            Else
                ' --- REGISTRO NUEVO ---
                 On Error Resume Next ' Habilitar manejo de error para la adición de fila
                 Set fila = tbl.ListRows.Add(AlwaysInsert:=True) ' Agregar nueva fila
                 If Err.Number <> 0 Then
                    MsgBox "Error Crítico al intentar agregar una nueva fila a la tabla '" & tbl.Name & "'." & vbCrLf & "Detalle: " & Err.Description & vbCrLf & "El proceso se detendrá.", vbCritical
                    Err.Clear
                    GoTo CleanupAndExit ' Ir a la salida limpia
                 End If
                 On Error GoTo ErrorHandler ' Restaurar manejo de errores general

                ' Rellenar las 4 columnas clave y las de conteo/fecha (con manejo de error individual)
                On Error Resume Next
                fila.Range(1, colIdxIdDeteccion).value = identificadorDeteccion
                If Err.Number <> 0 Then Debug.Print "Err escribiendo IdDeteccion: " & Err.Description: Err.Clear
                fila.Range(1, colIdxTipoOrigen).value = "Acunetix"
                 If Err.Number <> 0 Then Debug.Print "Err escribiendo TipoOrigen: " & Err.Description: Err.Clear
                fila.Range(1, colIdxIdVuln).value = csvName
                 If Err.Number <> 0 Then Debug.Print "Err escribiendo IdVuln: " & Err.Description: Err.Clear
                fila.Range(1, colIdxNomVuln).value = csvName
                 If Err.Number <> 0 Then Debug.Print "Err escribiendo NomVuln: " & Err.Description: Err.Clear
                fila.Range(1, colIdxConteo).value = 1 ' Iniciar conteo en 1
                 If Err.Number <> 0 Then Debug.Print "Err escribiendo Conteo: " & Err.Description: Err.Clear
                fila.Range(1, colIdxFecha).value = CDate(fechaActual) ' Intentar formato fecha
                If Err.Number <> 0 Then
                    Err.Clear
                    fila.Range(1, colIdxFecha).value = fechaActual ' Usar texto si falla
                     If Err.Number <> 0 Then Debug.Print "Err escribiendo Fecha: " & Err.Description: Err.Clear
                End If
                On Error GoTo ErrorHandler ' Restaurar manejo normal

                Set fila = Nothing ' Liberar memoria de la fila nueva
            End If

            Set filaExistente = Nothing ' Liberar referencia

SiguienteFilaCSV: ' Etiqueta para saltar a la siguiente iteración del bucle de filas CSV
        Next j ' Siguiente fila del CSV

CerrarYSaltar: ' Etiqueta para asegurar el cierre del CSV si se salta antes
        If Not wbCSV Is Nothing Then
             If wbCSV.Name = Mid(archivos(i), InStrRev(archivos(i), "\") + 1) Then ' Asegurarse que es el libro correcto
                wbCSV.Close False
             End If
        End If
        Set wsCSV = Nothing
        Set wbCSV = Nothing

SiguienteArchivo:
    Next i ' Siguiente archivo CSV

CleanupAndExit: ' Etiqueta para la salida limpia
    ' Reactivar cálculos y actualizaciones
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    ' Verificar si salimos por error o completado normal
    If Err.Number = 0 Then
       MsgBox "Proceso completado. Se verificaron duplicados y se actualizaron/agregaron registros en '" & tbl.Name & "'.", vbInformation
    End If
    ' Liberar objetos restantes (aunque ErrorHandler también lo hace)
    Set tbl = Nothing: Set ws = Nothing: Set wb = Nothing
    Set wbCSV = Nothing: Set wsCSV = Nothing
    Set fila = Nothing: Set filaExistente = Nothing: Set lrExisting = Nothing: Set t = Nothing
    Exit Sub ' Salir normalmente

ErrorHandler:
    ' Mostrar mensaje de error más detallado
    MsgBox "Ocurrió un error inesperado:" & vbCrLf & _
           "Número de Error: " & Err.Number & vbCrLf & _
           "Descripción: " & Err.Description & vbCrLf & _
           "Fuente: " & Err.Source & vbCrLf & _
           "Puede haber ocurrido en el archivo: " & IIf(i > 0 And i <= UBound(archivos), Mid(archivos(i), InStrRev(archivos(i), "\") + 1), "N/A") & _
           ", Fila CSV (aprox): " & j, _
           vbCritical, "Error en Macro"

    ' Intentar cerrar el CSV si aún está abierto
    If Not wbCSV Is Nothing Then
        On Error Resume Next ' Ignorar errores al intentar cerrar
        wbCSV.Close False
        On Error GoTo 0 ' Volver al manejo normal de errores (o a ninguno si no hay más código)
    End If

    ' Reactivar cálculos y actualizaciones en caso de error
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    ' Liberar objetos para evitar problemas
    Set tbl = Nothing: Set ws = Nothing: Set wb = Nothing
    Set wbCSV = Nothing: Set wsCSV = Nothing
    Set fila = Nothing: Set filaExistente = Nothing: Set lrExisting = Nothing: Set t = Nothing
    ' No usar Exit Sub aquí, permite que VBA termine limpiamente tras el error.

End Sub

    Sub CheckColumnExists(ByVal tblToCheck As ListObject, ByVal columnName As String, ByRef missingList As String, ByRef outIndex As Long)
        On Error Resume Next ' Habilitar manejo de error local para esta comprobación
        outIndex = 0 ' Reiniciar por si acaso
        outIndex = tblToCheck.ListColumns(columnName).Index
        If Err.Number <> 0 Then
            ' Si hay error, la columna no existe o el nombre es incorrecto
            If Len(missingList) > 0 Then missingList = missingList & ", " ' Añadir coma si ya hay elementos
            missingList = missingList & "'" & columnName & "'" ' Añadir el nombre de la columna faltante
            Err.Clear ' Limpiar el error para la siguiente comprobación
        End If
        On Error GoTo 0 ' Desactivar manejo de error local (vuelve al general o ninguno)
    End Sub

Sub CYB040_ResaltarFalsosPositivosEnVerde()
   Dim ws As Worksheet
    Dim tbl As ListObject
    Dim celda As Range
    Dim valoresTabla As Object
    Dim columnaIndex As Integer
    Dim encontrada As Boolean
    
    ' Crear un diccionario para almacenar los valores de la tabla
    Set valoresTabla = CreateObject("Scripting.Dictionary")
    
    ' Buscar la tabla en todas las hojas
    encontrada = False
    For Each ws In ThisWorkbook.Sheets
        On Error Resume Next
        Set tbl = ws.ListObjects("Tbl_falses_positives")
        On Error GoTo 0

        If Not tbl Is Nothing Then
            encontrada = True
            Exit For
        End If
    Next ws

    ' Si no se encuentra la tabla, mostrar mensaje y salir
    If Not encontrada Then
        MsgBox "No se encontr? la tabla 'Tbl_falses_positives' en ninguna hoja.", vbExclamation
        Exit Sub
    End If

    ' Obtener el índice de la columna "Vulnerability Name"
    On Error Resume Next
    columnaIndex = tbl.ListColumns("Vulnerability Name").Index
    On Error GoTo 0
    If columnaIndex = 0 Then
        MsgBox "La columna 'Vulnerability Name' no se encontr? en la tabla.", vbExclamation
        Exit Sub
    End If

    ' Guardar los valores de la columna de la tabla en el diccionario
    Dim celdaTabla As Range
    For Each celdaTabla In tbl.ListColumns(columnaIndex).DataBodyRange
        valoresTabla(celdaTabla.value) = True
    Next celdaTabla

    ' Recorrer las celdas seleccionadas y resaltar coincidencias
    Dim coincidencias As Boolean
    coincidencias = False
    For Each celda In Selection
        If valoresTabla.Exists(celda.value) Then
            celda.Interior.Color = RGB(0, 255, 0) ' Verde chill?n
            coincidencias = True
        End If
    Next celda

    ' Mensaje si hubo coincidencias o no
    If coincidencias Then
        MsgBox "Se han resaltado las celdas seleccionadas que coinciden con valores en 'Tbl_falses_positives'.", vbInformation
    Else
        MsgBox "No hay coincidencias en la tabla.", vbExclamation
    End If
End Sub

Sub CYB041_IrACatalogoVulnerabilidad()

    Dim wsOrigen As Worksheet, wsCatalogo As Worksheet
    Dim tblOrigen As ListObject, tblCatalogo As ListObject
    Dim rngCeldaActual As Range
    Dim idVulnerabilidad As Variant, tipoOrigen As String
    Dim colBusqueda As String, rngBusqueda As Range, celdaEncontrada As Range
    Dim dictColumnas As Object
    Dim nuevaFila As ListRow
    Dim respuesta As VbMsgBoxResult
    Dim colSourceDetection As Long, colLastEditedBy As Long, colLastUpdateDate As Long
    Dim fechaActual As String

    ' Definir las hojas
    Set wsOrigen = ActiveSheet
    Set wsCatalogo = ThisWorkbook.Sheets("Catalogo vulnerabilidades")
    
    ' Identificar la tabla en la hoja actual
    If wsOrigen.ListObjects.Count = 0 Then
        MsgBox "No se encontró una tabla en la hoja actual.", vbExclamation, "Error"
        Exit Sub
    End If
    Set tblOrigen = wsOrigen.ListObjects(1)
    
    ' Crear diccionario
    Set dictColumnas = CreateObject("Scripting.Dictionary")
    dictColumnas.Add "Nessus", "NessusPluginId"
    dictColumnas.Add "Invicti", "InvictiName"
    dictColumnas.Add "VulnerabilityManagerPlus", "VulnerabilityManagerPlusName"
    dictColumnas.Add "SonarQube", "SonarRuleId"
    dictColumnas.Add "DerScanner", "DerScannerName"
    dictColumnas.Add "Roslynator", "RoslynatorId"
    dictColumnas.Add "OWASPZAP", "OWASPZAPScanRuleId"
    dictColumnas.Add "Acunetix", "AcunetixName"
    dictColumnas.Add "OpenVas", "OpenVasNVTId"
    dictColumnas.Add "Nexpose", "NexposeName"
    dictColumnas.Add "InsightAppSec", "InsightAppSecInsightAppSec"
    dictColumnas.Add "Nmap", "NmapScriptName"
    dictColumnas.Add "Fortify", "FortifyName"
    dictColumnas.Add "Manual", "StandardVulnerabilityName"
    
    Set rngCeldaActual = ActiveCell
    If Intersect(rngCeldaActual, tblOrigen.DataBodyRange) Is Nothing Then
        MsgBox "Selecciona una celda dentro de la tabla de vulnerabilidades.", vbExclamation, "Error"
        Exit Sub
    End If
    
    tipoOrigen = tblOrigen.ListColumns("Tipo de origen").DataBodyRange.Cells(rngCeldaActual.Row - tblOrigen.DataBodyRange.Row + 1, 1).value
    idVulnerabilidad = tblOrigen.ListColumns("Identificador original de la vulnerabilidad").DataBodyRange.Cells(rngCeldaActual.Row - tblOrigen.DataBodyRange.Row + 1, 1).value
    
    If tipoOrigen = "" Or IsEmpty(idVulnerabilidad) Then
        MsgBox "Falta el Tipo de Origen o Identificador en la fila seleccionada.", vbExclamation, "Error"
        Exit Sub
    End If
    
    If Not dictColumnas.Exists(tipoOrigen) Then
        MsgBox "El tipo de origen '" & tipoOrigen & "' no tiene una columna asignada en el catálogo.", vbExclamation, "Error"
        Exit Sub
    End If
    
    colBusqueda = dictColumnas(tipoOrigen)
    
    On Error Resume Next
    Set tblCatalogo = wsCatalogo.ListObjects("Tbl_Catalogo_vulnerabilidades")
    On Error GoTo 0
    
    If tblCatalogo Is Nothing Then
        MsgBox "No se encontró la tabla 'Tbl_Catalogo_vulnerabilidades'.", vbExclamation, "Error"
        Exit Sub
    End If
    
    Set rngBusqueda = tblCatalogo.ListColumns(colBusqueda).DataBodyRange
    Set celdaEncontrada = rngBusqueda.Find(What:=idVulnerabilidad, LookAt:=xlWhole)
    
    If Not celdaEncontrada Is Nothing Then
        wsCatalogo.Activate
        celdaEncontrada.EntireRow.Select
        MsgBox "Registro encontrado. Se ha seleccionado la fila correspondiente en el catálogo.", vbInformation, "Éxito"
    Else
        respuesta = MsgBox("No se encontró el identificador en el catálogo. ¿Deseas agregarlo?", vbYesNo + vbQuestion, "Agregar nuevo registro")
        If respuesta = vbYes Then
            Set nuevaFila = tblCatalogo.ListRows.Add
            
            ' Insertar el ID en la columna correspondiente
            On Error Resume Next
            tblCatalogo.ListColumns(colBusqueda).DataBodyRange.Cells(nuevaFila.Index, 1).value = idVulnerabilidad
            On Error GoTo 0
            
            ' Rellenar columnas adicionales
            fechaActual = Format(Now, "dd/mm/yyyy")
            
            On Error Resume Next
            tblCatalogo.ListColumns("SourceDetection").DataBodyRange.Cells(nuevaFila.Index, 1).value = tipoOrigen
            tblCatalogo.ListColumns("LastEditedBy").DataBodyRange.Cells(nuevaFila.Index, 1).value = "Default System"
            tblCatalogo.ListColumns("LastUpdateDate").DataBodyRange.Cells(nuevaFila.Index, 1).value = fechaActual
            On Error GoTo 0
            
            ' *** Usar Application.Goto para seleccionar la fila sin activar la hoja ***
            Application.GoTo Reference:=tblCatalogo.ListRows(nuevaFila.Index).Range, Scroll:=True
            
            MsgBox "Nuevo registro agregado al catálogo. Se ha seleccionado la fila correspondiente.", vbInformation, "Éxito"
        End If
    End If
End Sub


Sub CYB042_Estandarizar()
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim rng As Range, cell As Range
    Dim dict As Object
    Dim colIndex As Object
    Dim key As String
    Dim i As Long, j As Long
    
    ' Definir la hoja activa
    Set ws = ActiveSheet
    Set dict = CreateObject("Scripting.Dictionary")
    Set colIndex = CreateObject("Scripting.Dictionary")
    
    ' Encontrar la última fila y última columna con datos
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Verificar si la columna StandardVulnerabilityName existe
    Dim stdCol As Integer
    stdCol = 0
    
    For i = 1 To lastCol
        If ws.Cells(1, i).value = "StandardVulnerabilityName" Then
            stdCol = i
        End If
        ' Guardar índices de columnas relevantes
        If Not IsEmpty(ws.Cells(1, i).value) Then
            colIndex(ws.Cells(1, i).value) = i
        End If
    Next i
    
    If stdCol = 0 Then
        MsgBox "La columna 'StandardVulnerabilityName' no existe.", vbExclamation
        Exit Sub
    End If
    
    ' Recorrer la tabla y agrupar datos
    For i = 2 To lastRow
        key = ws.Cells(i, stdCol).value
        If key <> "" Then
            If Not dict.Exists(key) Then
                dict.Add key, CreateObject("Scripting.Dictionary")
            End If
            
            ' Guardar valores no vacíos en cada columna relevante
            For Each colName In colIndex.Keys
                Dim colNum As Integer
                colNum = colIndex(colName)
                If ws.Cells(i, colNum).value <> "" Then
                    dict(key)(colName) = ws.Cells(i, colNum).value
                End If
            Next
        End If
    Next i
    
    ' Rellenar valores en base a los datos agrupados
    For i = 2 To lastRow
        key = ws.Cells(i, stdCol).value
        If key <> "" And dict.Exists(key) Then
            For Each colName In colIndex.Keys
                colNum = colIndex(colName)
                If ws.Cells(i, colNum).value = "" And dict(key).Exists(colName) Then
                    ws.Cells(i, colNum).value = dict(key)(colName)
                End If
            Next
        End If
    Next i
    
    MsgBox "Estandarizaci?n completada.", vbInformation
End Sub




Sub CYB043_AplicarFormatoCondicional()
    Dim selectedRange As Range

    ' Verificar si hay celdas seleccionadas
    On Error Resume Next
    Set selectedRange = Selection.SpecialCells(xlCellTypeConstants)
    On Error GoTo 0

    ' Salir si no hay celdas seleccionadas
    If selectedRange Is Nothing Then
        MsgBox "No hay celdas seleccionadas."
        Exit Sub
    End If

    ' Aplicar formato condicional según el contenido de las celdas seleccionadas
    With selectedRange
        .FormatConditions.Add Type:=xlTextString, String:="CRÍTICA", TextOperator:=xlContains
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Font
            .Color = RGB(255, 255, 255) ' Blanco
        End With
        With .FormatConditions(1).Interior
            .Color = RGB(112, 48, 160) ' #7030A0
        End With

        .FormatConditions.Add Type:=xlTextString, String:="ALTA", TextOperator:=xlContains
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Font
            .Color = RGB(255, 255, 255) ' Blanco
        End With
        With .FormatConditions(1).Interior
            .Color = RGB(255, 0, 0) ' #FF0000
        End With

        .FormatConditions.Add Type:=xlTextString, String:="MEDIA", TextOperator:=xlContains
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Font
            .Color = RGB(0, 0, 0) ' Negro
        End With
        With .FormatConditions(1).Interior
            .Color = RGB(255, 255, 0) ' #FFFF00
        End With

        .FormatConditions.Add Type:=xlTextString, String:="BAJA", TextOperator:=xlContains
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Font
            .Color = RGB(255, 255, 255) ' Blanco
        End With
        With .FormatConditions(1).Interior
            .Color = RGB(0, 176, 80) ' #00B050
        End With

        .FormatConditions.Add Type:=xlTextString, String:="INFORMATIVA", TextOperator:=xlContains
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Font
            .Color = RGB(0, 0, 0) ' Negro
        End With
        With .FormatConditions(1).Interior
            .Color = RGB(231, 230, 230) ' #E7E6E6
        End With
    End With
End Sub

Sub CYB061_LLM_llama3_2_1b()
     Dim http As Object
    Dim JSONBody As String
    Dim response As String
    Dim Vulnerabilidad As String
    Dim extractedResponse As String
    Dim cell As Range
    
    ' Verificar si hay celdas seleccionadas
    If Selection.Cells.Count = 0 Then
        MsgBox "Seleccione al menos una celda con una vulnerabilidad antes de ejecutar la macro.", vbExclamation, "Error"
        Exit Sub
    End If
    
    ' Crear objeto HTTP
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' Recorrer cada celda en la selecci?n
    For Each cell In Selection
        ' Verificar si la celda no est? vacía
        If Not IsEmpty(cell.value) Then
            Vulnerabilidad = cell.value
            
            ' Construcci?n del prompt
            Dim prompt As String
            prompt = ConstruirPrompt(Vulnerabilidad)
            
            ' Crear el cuerpo del JSON
            JSONBody = "{""model"": ""llama3.2:1b"", ""prompt"": """ & Replace(prompt, """", "\""") & """, ""stream"": false}"
            
            ' Enviar la solicitud HTTP
            With http
                .Open "POST", "http://localhost:11434/api/generate", False
                .setRequestHeader "Content-Type", "application/json"
                .Send JSONBody
                response = .responseText
            End With
            
            ' Extraer solo el valor de "response.response"
            extractedResponse = ExtraerRespuesta(response)
            
            ' Asignar la respuesta a la celda correspondiente
            cell.value = extractedResponse
        End If
    Next cell
    
    ' Liberar objeto HTTP
    Set http = Nothing
End Sub


Sub CYB060_LLLM_deepseek_r1_1_5b()
    Dim http As Object
    Dim JSONBody As String
    Dim response As String
    Dim Vulnerabilidad As String
    Dim extractedResponse As String
    Dim cell As Range
    
    ' Verificar si hay celdas seleccionadas
    If Selection.Cells.Count = 0 Then
        MsgBox "Seleccione al menos una celda con una vulnerabilidad antes de ejecutar la macro.", vbExclamation, "Error"
        Exit Sub
    End If
    
    ' Crear objeto HTTP
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' Recorrer cada celda en la selecci?n
    For Each cell In Selection
        ' Verificar si la celda no est? vacía
        If Not IsEmpty(cell.value) Then
            Vulnerabilidad = cell.value
            
            ' Construcci?n del prompt
            Dim prompt As String
            prompt = ConstruirPrompt(Vulnerabilidad)
            
            ' Crear el cuerpo del JSON
            JSONBody = "{""model"": ""deepseek-r1:1.5b"", ""prompt"": """ & Replace(prompt, """", "\""") & """, ""stream"": false}"
            
            ' Enviar la solicitud HTTP
            With http
                .Open "POST", "http://localhost:11434/api/generate", False
                .setRequestHeader "Content-Type", "application/json"
                .Send JSONBody
                response = .responseText
            End With
            
            ' Extraer solo el valor de "response.response"
            extractedResponse = ExtraerRespuesta(response)
            
            ' Asignar la respuesta a la celda correspondiente
            cell.value = extractedResponse
        End If
    Next cell
    
    ' Liberar objeto HTTP
    Set http = Nothing
End Sub




' Funci?n para construir el prompt de forma m?s clara y estructurada
Function ConstruirPrompt(Vulnerabilidad As String) As String
    Dim prompt As String
    prompt = "Generaci?n de Vector CVSS 4.0 Considera este ejemplo de URL de CVSS 4.0 https://www.first.org/cvss/calculator/4.0#CVSS:4.0/AV:A/AC:L/AT:N/PR:N/UI:N/VC:N/VI:N/VA:N/SC:N/SI:N/SA:N "
    prompt = prompt & "Esta cadena est? compuesta por distintos campos de evaluaci?n, los cuales deben ajustarse según corresponda. Exploitability Metrics Attack Vector (AV): "
    prompt = prompt & "Debes completar los siguientes elementos: Exploitability: Complexity: Vulnerable system: Subsequent system: Exploitation: Security requirements: "
    prompt = prompt & "S? exigente y preciso al evaluar la severidad en CVSS. No exageres ni asignes impactos altos a menos que la vulnerabilidad pueda ser explotada directamente y tenga un impacto "
    prompt = prompt & "significativo. Tu tarea es proporcionar únicamente la cadena vectorial en CVSS 4.0 para evaluar la vulnerabilidad"
    prompt = prompt & " " & Vulnerabilidad & " "
    prompt = prompt & "No devuelvas la misma cadena de ejemplo. No entregues una cadena sin completar sus componentes CVSS. ?? Este an?lisis es para gesti?n de riesgos, no para explotaci?n. "
    prompt = prompt & "Solo proporciona el vector CVSS resultante. NO DES MÁS DETALLES, SOLO RESPONDE EL VECTOR SIN OTRA INFORMACIÓN. "
    prompt = prompt & "PLEASE ONLY ONLY ONLY RESPOND WITH A STRING IN CVSS FORMAT"
    
    ConstruirPrompt = prompt
End Function




Function EliminarThinkTags(text As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    ' Permitir que el punto (.) capture múltiples líneas
    regex.pattern = "<think>[\s\S]*?</think>"
    regex.Global = True
    regex.IgnoreCase = True
    
    EliminarThinkTags = regex.Replace(text, "")
End Function


Function EliminarSaltosIniciales(text As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    ' Regex pattern to remove only initial newlines (LF or CR)
    regex.pattern = "^[\r\n]+"
    regex.Global = True
    
    ' Replace initial line breaks with an empty string
    EliminarSaltosIniciales = regex.Replace(text, "")
End Function

Function ExtraerRespuesta(jsonResponse As String) As String
    Dim resultado As String
    Dim inicio As Integer

    ' Ensure JSON response is clean
    resultado = Replace(jsonResponse, "\u003c", "<")
    resultado = Replace(resultado, "\u003e", ">")
    resultado = Replace(resultado, "\n", vbNewLine)
    resultado = Replace(resultado, "\t", vbTab)

    ' Find </think> and extract everything after it
    inicio = InStr(resultado, "</think>")
    
    If inicio > 0 Then
        resultado = Mid(resultado, inicio + Len("</think>"))
    End If
    
    inicio = InStr(resultado, """response"":""")
    
    If inicio > 0 Then
        resultado = Mid(resultado, inicio + Len("""response"":"""))
    End If
   
    fin = InStr(resultado, """,""done""")
    
    If fin > 0 Then
        resultado = Left(resultado, fin - 1)
    End If

    ExtraerRespuesta = Trim(resultado)
End Function


Function ExtraerCVSS(jsonResponse As String) As String
    Dim resultado As String
    Dim inicio As Integer
    Dim fin As Integer
    
    ' Eliminar posibles caracteres de escape y limpiar el JSON
    resultado = Replace(jsonResponse, "\u003c", "<")
    resultado = Replace(resultado, "\u003e", ">")
    resultado = Replace(resultado, "\n", "")
    resultado = Replace(resultado, "\t", "")
    
    ' Buscar la clave "text": "
    inicio = InStr(resultado, """text"": """)
    
    If inicio > 0 Then
        ' Extraer el texto despu?s de "text": "
        resultado = Mid(resultado, inicio + Len("""text"": """))
        
        ' Buscar la posici?n final antes del cierre de comillas
        fin = InStr(resultado, """")
        If fin > 0 Then
            resultado = Left(resultado, fin - 1)
        End If
    Else
        resultado = "No se encontr? CVSS"
    End If

    ' Retornar el CVSS extraído
    ExtraerCVSS = Trim(resultado)
End Function



Sub ObtenerRespuestasGeminiCVSS4()
    Dim cell As Range
    Dim http As Object
    Dim json As Object
    Dim apiUrl As String
    Dim apiKey As String
    Dim requestData As String
    Dim responseText As String
    Dim answerID As String
    
    ' Clave de API de Gemini (reempl?zala con la tuya) AIzaSyBbd_upGJ2JzdsmWSzNBvSr3mXiPo9h4bs  AIzaSyADfixgVHPBXyY60ivLUYo3rCJTQtZ_M7g
    apiKey = "AIzaSyBbd_upGJ2JzdsmWSzNBvSr3mXiPo9h4bs"
    apiUrl = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash-lite:generateContent?key=" & apiKey

       ' Crear el objeto HTTP
    Set http = CreateObject("MSXML2.XMLHTTP")

    ' Iterar sobre cada celda seleccionada
    For Each cell In Selection
     
     Dim promptvalue As String
     promptvalue = ConstruirPrompt(cell.value)
    
    
        ' Construir el prompt con la pregunta de la celda
        requestData = "{""contents"": [{""parts"": [{""text"": """ & promptvalue & """}]}]}"


        ' Enviar la solicitud HTTP
        With http
            .Open "POST", apiUrl, False
            .setRequestHeader "Content-Type", "application/json"
            .Send requestData
        End With

        ' Procesar la respuesta
        If http.Status = 200 Then
            responseText = http.responseText
            Debug.Print "Response: " & responseText ' Imprime la respuesta completa en la ventana inmediata

            ' Intentar analizar JSON
            On Error Resume Next
            Set json = JsonConverter.ParseJson(responseText)
            On Error GoTo 0

            ' Validar JSON
            If Not json Is Nothing Then
                If json.Exists("candidates") And json("candidates").Count > 0 Then
                    If json("candidates")(0).Exists("content") And json("candidates")(0)("content").Exists("parts") Then
                        If json("candidates")(0)("content")("parts").Count > 0 Then
                            answerID = json("candidates")(0)("content")("parts")(0)("text")
                        Else
                            answerID = "No CVSS data"
                        End If
                    Else
                        answerID = "Invalid response format"
                    End If
                Else
                    answerID = "No candidates found"
                End If
            Else
                answerID = "Error parsing JSON"
            End If
        Else
            answerID = "HTTP Error: " & http.Status
            Debug.Print "Response: " & responseText ' Imprime la respuesta completa en la ventana inmediata
        End If

        ' Colocar la respuesta en la celda adyacente
        cell.Offset(0, 1).value = ExtraerCVSS(responseText)
    Next cell

    ' Liberar objetos
    Set http = Nothing
    Set json = Nothing

    MsgBox "Procesamiento completado.", vbInformation
End Sub




' Macro Modificada 1: Descripción General de Vulnerabilidad
Sub CYB068_PrepararPromptDesdeSeleccion_DescripcionVuln_General()
    Dim celda As Range
    Dim listaVulnerabilidades As String
    Dim prompt As String

    ' Inicializar la variable que almacenará las vulnerabilidades
    listaVulnerabilidades = ""

    ' Verificar si la selección es válida
    If TypeName(Selection) <> "Range" Then
        MsgBox "Por favor, selecciona un rango de celdas.", vbExclamation, "Selección Inválida"
        Exit Sub
    End If

    ' Recorrer las celdas seleccionadas y acumular los valores con saltos de línea y comas
    For Each celda In Selection
        If Not IsEmpty(celda.value) Then
            If listaVulnerabilidades <> "" Then
                listaVulnerabilidades = listaVulnerabilidades & "," & vbCrLf
            End If
            listaVulnerabilidades = listaVulnerabilidades & Trim(CStr(celda.value)) ' Usar Trim y CStr para robustez
        End If
    Next celda

    ' Construir el prompt final
    If listaVulnerabilidades <> "" Then
        prompt = "Redacta un párrafo técnico breve y conciso que describa la vulnerabilidad detectada, comenzando con la frase: -El sistema...-. Explica en qué consiste la debilidad de seguridad de manera técnica. No incluyas escenarios de explotación, ya que eso corresponde a otro campo. No describas cómo se explota, solo en qué consiste el problema. No menciones el nombre exacto de la vulnerabilidad; utiliza expresiones similares. Vulnerabilidades a describir en formato tabla: "
        prompt = prompt & " " & vbCrLf & Chr(10) & listaVulnerabilidades & vbCrLf & Chr(10)
        ' Contexto Generalizado
        prompt = prompt & "Contexto de análisis: Análisis de vulnerabilidades realizado en el entorno evaluado. Vulnerabilidad detectada mediante escaneo e interacciones en el sistema bajo revisión."

        ' Copiar el prompt generado al portapapeles
        CopiarAlPortapapeles prompt

        ' Mostrar un mensaje informativo
        MsgBox "El prompt para descripción de vulnerabilidad ha sido copiado al portapapeles.", vbInformation, "Prompt Generado"
    Else
        MsgBox "No se encontraron valores en la selección.", vbExclamation, "Sin Datos"
    End If
End Sub

' Macro Modificada 2: Amenaza General de Vulnerabilidad
Sub CYB069_PrepararPromptDesdeSeleccion_AmenazaVuln_General()
    Dim celda As Range
    Dim listaVulnerabilidades As String
    Dim prompt As String

    ' Inicializar la variable que almacenará las vulnerabilidades
    listaVulnerabilidades = ""

    ' Verificar si la selección es válida
    If TypeName(Selection) <> "Range" Then
        MsgBox "Por favor, selecciona un rango de celdas.", vbExclamation, "Selección Inválida"
        Exit Sub
    End If

    ' Recorrer las celdas seleccionadas y acumular los valores con saltos de línea y comas
    For Each celda In Selection
        If Not IsEmpty(celda.value) Then
            If listaVulnerabilidades <> "" Then
                listaVulnerabilidades = listaVulnerabilidades & "," & vbCrLf
            End If
            listaVulnerabilidades = listaVulnerabilidades & Trim(CStr(celda.value)) ' Usar Trim y CStr para robustez
        End If
    Next celda

    ' Construir el prompt final
    If listaVulnerabilidades <> "" Then
        prompt = "Hola, por favor considera el siguiente ejemplo: Un atacante podría (obtener, realizar, ejecutar, visualizar, identificar, listar)... Algunos posibles vectores de ataque adicionales asociados con esta amenaza incluyen (pueden ser algunos de los siguientes, pero no necesariamente todos): • Malware: Un malware diseñado para automatizar intentos de fuerza bruta podría explotar la vulnerabilidad para... • Usuario malintencionado: Un usuario dentro del entorno con conocimiento de la vulnerabilidad podría aprovecharla para... • Personal interno: Un empleado con acceso y conocimientos técnicos podría, intencionalmente o por error,... • Delincuente cibernético: Un atacante externo en busca de vulnerabilidades podría intentar explotar esta debilidad para... Instrucciones adicionales: "
        prompt = prompt & "1. Pregunta si el sistema es interno o accesible externamente para determinar los vectores de ataque más relevantes, ya que no todos aplican en todos los casos. " ' Mantenido pero generalizado levemente
        prompt = prompt & "2. Redacta una descripción de la amenaza que incluya posibles escenarios de ataque, comenzando con la frase: -Un atacante podría...-. "
        prompt = prompt & "3. No es necesario proporcionar un ejemplo para cada vector de ataque. Selecciona el más realista o probable. "
        prompt = prompt & "4. En las viñetas de vectores de ataque, menciona la probabilidad de ocurrencia, incluso si es baja. "
        ' Contexto Generalizado
        prompt = prompt & "5. El contexto es un análisis de vulnerabilidades de infraestructura. Vulnerabilidad detectada mediante escaneo e interacciones. Formato de respuesta: • Responde en una tabla de dos columnas. • Para cada vulnerabilidad, redacta un párrafo descriptivo en la primera columna (mínimo 75 palabras). • En la segunda columna, lista los vectores de ataque con viñetas (usando guiones - ). • No uses HTML, solo texto plano. Ejemplo de estructura: Descripción de la amenaza   Vectores de ataque Un atacante podría explotar esta vulnerabilidad para acceder información"
        prompt = prompt & " confidencial. Esta amenaza es particularmente crítica en sistemas donde los controles de seguridad son menos estrictos."
        prompt = prompt & " Un escenario probable incluye...   - Malware: Un malware podría ser utilizado para... (probabilidad media). - Usuario malintencionado: Un empleado con acceso podría... (probabilidad baja). - Delincuente cibernético: Un atacante externo podría... (probabilidad alta). ES MUY IMPORTANTE QUE PARA LOS VECTORES DE ATAQUE DE LA AMENAZA USES GUIONES MEDIOS COMO VIÑETAS DENTRO DE LAS CELDAS. "
        ' Eliminada la mención específica de detección, reemplazada por generalidad
        prompt = prompt & "Explica los escenarios que consideres necesarios según la naturaleza de la vulnerabilidad." & vbCrLf & Chr(10)
        prompt = prompt & "SOLO DOS columnas: Nombre (o descripción breve de la vulnerabilidad) y Amenaza/Vectores." ' Clarificado el nombre de la primera columna
        prompt = prompt & " " & vbCrLf & Chr(10) & listaVulnerabilidades & vbCrLf & Chr(10)

        ' Copiar el prompt generado al portapapeles
        CopiarAlPortapapeles prompt

        ' Mostrar un mensaje informativo
        MsgBox "El prompt para amenaza de vulnerabilidad ha sido copiado al portapapeles.", vbInformation, "Prompt Generado"
    Else
        MsgBox "No se encontraron valores en la selección.", vbExclamation, "Sin Datos"
    End If
End Sub

' Macro Modificada 3: Propuesta General de Remediación
Sub CYB070_PrepararPromptDesdeSeleccion_PropuestaRemediacionVuln_General()
    Dim celda As Range
    Dim listaVulnerabilidades As String
    Dim prompt As String

    ' Inicializar la variable que almacenará las vulnerabilidades
    listaVulnerabilidades = ""

    ' Verificar si la selección es válida
    If TypeName(Selection) <> "Range" Then
        MsgBox "Por favor, selecciona un rango de celdas.", vbExclamation, "Selección Inválida"
        Exit Sub
    End If

    ' Recorrer las celdas seleccionadas y acumular los valores con saltos de línea y comas
    For Each celda In Selection
        If Not IsEmpty(celda.value) Then
            If listaVulnerabilidades <> "" Then
                listaVulnerabilidades = listaVulnerabilidades & "," & vbCrLf
            End If
            listaVulnerabilidades = listaVulnerabilidades & Trim(CStr(celda.value)) ' Usar Trim y CStr para robustez
        End If
    Next celda

    ' Construir el prompt final
    If listaVulnerabilidades <> "" Then
        prompt = "Hola, por favor Redacta como un pentester un párrafo técnico de propuesta de remediación que comience con la frase: -Se recomienda...-. Incluye tantos detalles puntuales para la remediación como sea posible, por ejemplo: nombres de soluciones que funcionan como protección, controles de seguridad específicos, dispositivos, configuraciones, buenas prácticas. Menciona de manera puntual qué se podría hacer para que el encargado del sistema o activo pueda saber cómo remediar. La respuesta debe tener SOLO un párrafo breve de introducción por vulnerabilidad y luego viñetas (usando guiones -) para los puntos de la propuesta de remediación. Responde para la siguiente lista de vulnerabilidades en FORMATO TABLA DE DOS COLUMNAS. SIEMPRE COMIENZA CON -Se recomienda...- TEXTO AMPLIO (más de 80 palabras por remediación), aplicable a diversos casos, mencionando tecnologías, lenguajes o frameworks si aplica."
        ' Eliminada la mención específica de detección. El contexto es implícito por las vulnerabilidades listadas.
        prompt = prompt & " Menciona solo soluciones y prácticas corporativas/profesionales."
        prompt = prompt & " Solo dos columnas: Nombre (o descripción breve de la vulnerabilidad) y Propuesta de Remediación." ' Clarificado el nombre de la primera columna
        prompt = prompt & " " & vbCrLf & Chr(10) & listaVulnerabilidades & vbCrLf & Chr(10)

        ' Copiar el prompt generado al portapapeles
        CopiarAlPortapapeles prompt

        ' Mostrar un mensaje informativo
        MsgBox "El prompt para propuesta de remediación ha sido copiado al portapapeles.", vbInformation, "Prompt Generado"
    Else
        MsgBox "No se encontraron valores en la selección.", vbExclamation, "Sin Datos"
    End If
End Sub

Sub CYB071_PrepararPromptDesdeSeleccion_DescripcionVuln_EnVPN()
    Dim celda As Range
    Dim listaVulnerabilidades As String
    Dim prompt As String
    
    ' Inicializar la variable que almacenar? las vulnerabilidades
    listaVulnerabilidades = ""
    
    ' Recorrer las celdas seleccionadas y acumular los valores con saltos de línea y comas
    For Each celda In Selection
        If Not IsEmpty(celda.value) Then
            If listaVulnerabilidades <> "" Then
                listaVulnerabilidades = listaVulnerabilidades & "," & vbCrLf
            End If
            listaVulnerabilidades = listaVulnerabilidades & celda.value
        End If
    Next celda
    
    ' Construir el prompt final
    If listaVulnerabilidades <> "" Then
        prompt = "Redacta un p?rrafo t?cnico breve y conciso que describa la vulnerabilidad detectada, comenzando con la frase: -El sistema...-. Explica en qu? consiste la debilidad de seguridad de manera t?cnica. No incluyas escenarios de explotaci?n, ya que eso corresponde a otro campo. No describas c?mo se explota, solo en qu? consiste el problema. No menciones el nombre exacto de la vulnerabilidad; utiliza expresiones similares. Vulnerabilidades a describi en formato tabla: "
        prompt = prompt & " " & vbCrLf & Chr(10) & listaVulnerabilidades & vbCrLf & Chr(10)
        prompt = prompt & "Contexto de an?lisis: An?lisis de vulnerabilidades de infraestructura a partir conexion a red VPN. Vulnerabilidad detectada mediante escaneo e interacciones desde red privada VPN."

        
        ' Copiar el prompt generado al portapapeles
        CopiarAlPortapapeles prompt
        
        ' Mostrar un mensaje informativo
        MsgBox "El prompt ha sido copiado al portapapeles.", vbInformation, "Prompt Generado"
    Else
        MsgBox "No se encontraron valores en la selecci?n.", vbExclamation, "Sin Datos"
    End If
End Sub



Sub CYB072_PrepararPromptDesdeSeleccion_AmenazaVuln_EnVPN()
    Dim celda As Range
    Dim listaVulnerabilidades As String
    Dim prompt As String
    
    ' Inicializar la variable que almacenar? las vulnerabilidades
    listaVulnerabilidades = ""
    
    ' Recorrer las celdas seleccionadas y acumular los valores con saltos de línea y comas
    For Each celda In Selection
        If Not IsEmpty(celda.value) Then
            If listaVulnerabilidades <> "" Then
                listaVulnerabilidades = listaVulnerabilidades & "," & vbCrLf
            End If
            listaVulnerabilidades = listaVulnerabilidades & celda.value
        End If
    Next celda
    
    ' Construir el prompt final
    If listaVulnerabilidades <> "" Then
        ' prompt = prompt & ""
        prompt = "Hola, por favor considera el siguiente ejemplo: Un atacante podría (obtener, realizar, ejecutar, visualizar, identificar, listar)... Algunos posibles vectores de ataque adicionales asociados con esta amenaza incluyen (pueden ser algunos de los siguientes, pero no necesariamente todos): •  Malware: Un malware diseñado para automatizar intentos de fuerza bruta podría explotar la vulnerabilidad para... •    Usuario malintencionado: Un usuario dentro de la red local con conocimiento de la vulnerabilidad podría aprovecharla para... •    Personal interno: Un empleado con acceso y conocimientos t?cnicos podría, intencionalmente o por error,... •  Delincuente cibern?tico: Un atacante externo en busca de vulnerabilidades podría intentar explotar esta debilidad para... Instrucciones adicionales: 1.   Pregunta si el sistema es interno o externo para determinar los vectores de ataque m?s relevantes, ya que no todos aplican en todos los casos. "
        prompt = prompt & "2.   Redacta una descripci?n de la amenaza que incluya posibles escenarios de ataque, comenzando con la frase: -Un atacante podría...-. 3.    No es necesario proporcionar un ejemplo para cada vector de ataque. Selecciona el m?s realista o probable. 4.   En las viñetas de vectores de ataque, menciona la probabilidad de ocurrencia, incluso si es baja. 5.    El contexto es un an?lisis de vulnerabilidades de infraestructura desde una red privada, específicamente: -Vulnerabilidad detectada mediante escaneo e interacciones desde red privada-. Formato de respuesta: •    Responde en una tabla de dos columnas. •    Para cada vulnerabilidad, redacta un p?rrafo descriptivo en la primera columna (mínimo 75 palabras). •  En la segunda columna, lista los vectores de ataque con viñetas (usando guiones  - ). • No uses HTML, solo texto plano. Ejemplo de estructura: Descripci?n de la amenaza    Vectores de ataque Un atacante podría explotar esta vulnerabilidad para acceder a inf _
maci?n"
        prompt = prompt & "confidencial Esta amenaza es particularmente crítica en sistemas internos donde los controles de seguridad son menos estrictos."
        prompt = prompt & "Un escenario probable incluye...    - Malware: Un malware podría ser utilizado para... (probabilidad media). - Usuario malintencionado: Un empleado con acceso podría... (probabilidad baja). - Delincuente cibern?tico: Un atacante externo podría... (probabilidad alta). ES MUY IMPORANTE QU EPAR ALOS VECTORI DE ATQUE DE LA MENAZA USSES GUIONE S MEDIOS COMO VIÑETAS DENTRO DE ALS CELDAS"
        prompt = prompt & "Se detecto mediante VPN red privada, pero explica los escenarios que consideres necesarios" & vbCrLf & Chr(10)
        prompt = prompt & "SOLO DOS columnas, nombre y amenaza"
        prompt = prompt & " " & vbCrLf & Chr(10) & listaVulnerabilidades & vbCrLf & Chr(10)
          
        ' Copiar el prompt generado al portapapeles
        CopiarAlPortapapeles prompt
        
        ' Mostrar un mensaje informativo
        MsgBox "El prompt ha sido copiado al portapapeles.", vbInformation, "Prompt Generado"
    Else
        MsgBox "No se encontraron valores en la selecci?n.", vbExclamation, "Sin Datos"
    End If
End Sub


Sub CYB073_PrepararPromptDesdeSeleccion_PropuestaRemediacionVuln_EnVPN()
    Dim celda As Range
    Dim listaVulnerabilidades As String
    Dim prompt As String
    
    ' Inicializar la variable que almacenar? las vulnerabilidades
    listaVulnerabilidades = ""
    
    ' Recorrer las celdas seleccionadas y acumular los valores con saltos de línea y comas
    For Each celda In Selection
        If Not IsEmpty(celda.value) Then
            If listaVulnerabilidades <> "" Then
                listaVulnerabilidades = listaVulnerabilidades & "," & vbCrLf
            End If
            listaVulnerabilidades = listaVulnerabilidades & celda.value
        End If
    Next celda
    
    ' Construir el prompt final
    If listaVulnerabilidades <> "" Then
        ' prompt = prompt & ""
        prompt = "Hola, por favor Redacta como un pentester un p?rrafo t?cnico de propuesta de remediaci?n que comience con la frase: -Se recomienda...-. Incluye tantos detalles puntuales par remaicion por jemplo nombre de soluciones que funeriona como poroteccon o controles de sguridad, disoitivos, practicas, menciona de manera m?s puntual que se pod´ria hacer para que le encargado del sistema, activo, pueda saber como remediari La respueste debe tener SOLO parrafo breve introducci?n y luego viñetas para lso putnos de propueta de remedaicion Responde para la siguientel lista de vulnerabilidades en FORMATO TABLA DOS COLUMNAS, , SIEMPRE COMIENZAX CON -Se recomienda...- TEXTO AMPLIDO MAS DE 80 palabras, atalmente aplicable a muchos casos, lengauje so framworks"
           prompt = prompt & "Se detecto mediante VPN red privada, pero explica los escenarios que consideres necesarios" & vbCrLf & Chr(10)
          prompt = prompt & "Menciona solo soluciones corporativas"
          prompt = prompt & "Solo dos columnas, nombre y propuesat remediacion"
        prompt = prompt & " " & vbCrLf & Chr(10) & listaVulnerabilidades & vbCrLf & Chr(10)
          
        ' Copiar el prompt generado al portapapeles
        CopiarAlPortapapeles prompt
        
        ' Mostrar un mensaje informativo
        MsgBox "El prompt ha sido copiado al portapapeles.", vbInformation, "Prompt Generado"
    Else
        MsgBox "No se encontraron valores en la selecci?n.", vbExclamation, "Sin Datos"
    End If
End Sub


Sub CYB074_PrepararPromptDesdeSeleccion_DescripcionVuln_EnRedPrivada()
    Dim celda As Range
    Dim listaVulnerabilidades As String
    Dim prompt As String
    
    ' Inicializar la variable que almacenar? las vulnerabilidades
    listaVulnerabilidades = ""
    
    ' Recorrer las celdas seleccionadas y acumular los valores con saltos de línea y comas
    For Each celda In Selection
        If Not IsEmpty(celda.value) Then
            If listaVulnerabilidades <> "" Then
                listaVulnerabilidades = listaVulnerabilidades & "," & vbCrLf
            End If
            listaVulnerabilidades = listaVulnerabilidades & celda.value
        End If
    Next celda
    
    ' Construir el prompt final
    If listaVulnerabilidades <> "" Then
        prompt = "Redacta un p?rrafo t?cnico breve y conciso que describa la vulnerabilidad detectada, comenzando con la frase: -El sistema...-. Explica en qu? consiste la debilidad de seguridad de manera t?cnica. No incluyas escenarios de explotaci?n, ya que eso corresponde a otro campo. No describas c?mo se explota, solo en qu? consiste el problema. No menciones el nombre exacto de la vulnerabilidad; utiliza expresiones similares. Vulnerabilidades a describi en formato tabla: "
        prompt = prompt & " " & vbCrLf & Chr(10) & listaVulnerabilidades & vbCrLf & Chr(10)
        prompt = prompt & "Contexto de an?lisis: An?lisis de vulnerabilidades de infraestructura a partir conexion a red en red privada en sitio. Vulnerabilidad detectada mediante escaneo e interacciones desde red privada en red privada en sitio."

        
        ' Copiar el prompt generado al portapapeles
        CopiarAlPortapapeles prompt
        
        ' Mostrar un mensaje informativo
        MsgBox "El prompt ha sido copiado al portapapeles.", vbInformation, "Prompt Generado"
    Else
        MsgBox "No se encontraron valores en la selecci?n.", vbExclamation, "Sin Datos"
    End If
End Sub



Sub CYB075_PrepararPromptDesdeSeleccion_AmenazaVuln_EnRedPrivada()
    Dim celda As Range
    Dim listaVulnerabilidades As String
    Dim prompt As String
    
    ' Inicializar la variable que almacenar? las vulnerabilidades
    listaVulnerabilidades = ""
    
    ' Recorrer las celdas seleccionadas y acumular los valores con saltos de línea y comas
    For Each celda In Selection
        If Not IsEmpty(celda.value) Then
            If listaVulnerabilidades <> "" Then
                listaVulnerabilidades = listaVulnerabilidades & "," & vbCrLf
            End If
            listaVulnerabilidades = listaVulnerabilidades & celda.value
        End If
    Next celda
    
    ' Construir el prompt final
    If listaVulnerabilidades <> "" Then
        ' prompt = prompt & ""
        prompt = "Hola, por favor considera el siguiente ejemplo: Un atacante podría (obtener, realizar, ejecutar, visualizar, identificar, listar)... Algunos posibles vectores de ataque adicionales asociados con esta amenaza incluyen (pueden ser algunos de los siguientes, pero no necesariamente todos): •  Malware: Un malware diseñado para automatizar intentos de fuerza bruta podría explotar la vulnerabilidad para... •    Usuario malintencionado: Un usuario dentro de la red local con conocimiento de la vulnerabilidad podría aprovecharla para... •    Personal interno: Un empleado con acceso y conocimientos t?cnicos podría, intencionalmente o por error,... •  Delincuente cibern?tico: Un atacante externo en busca de vulnerabilidades podría intentar explotar esta debilidad para... Instrucciones adicionales: 1.   Pregunta si el sistema es interno o externo para determinar los vectores de ataque m?s relevantes, ya que no todos aplican en todos los casos. "
        prompt = prompt & "2.   Redacta una descripci?n de la amenaza que incluya posibles escenarios de ataque, comenzando con la frase: -Un atacante podría...-. 3.    No es necesario proporcionar un ejemplo para cada vector de ataque. Selecciona el m?s realista o probable. 4.   En las viñetas de vectores de ataque, menciona la probabilidad de ocurrencia, incluso si es baja. 5.    El contexto es un an?lisis de vulnerabilidades de infraestructura desde una red privada, específicamente: -Vulnerabilidad detectada mediante escaneo e interacciones desde red privada-. Formato de respuesta: •    Responde en una tabla de dos columnas. •    Para cada vulnerabilidad, redacta un p?rrafo descriptivo en la primera columna (mínimo 75 palabras). •  En la segunda columna, lista los vectores de ataque con viñetas (usando guiones  - ). • No uses HTML, solo texto plano. Ejemplo de estructura: Descripci?n de la amenaza    Vectores de ataque Un atacante podría explotar esta vulnerabilidad para acceder a inf _
maci?n"
        prompt = prompt & "confidencial Esta amenaza es particularmente crítica en sistemas internos donde los controles de seguridad son menos estrictos."
        prompt = prompt & "Un escenario probable incluye...    - Malware: Un malware podría ser utilizado para... (probabilidad media). - Usuario malintencionado: Un empleado con acceso podría... (probabilidad baja). - Delincuente cibern?tico: Un atacante externo podría... (probabilidad alta). ES MUY IMPORANTE QU EPAR ALOS VECTORI DE ATQUE DE LA MENAZA USSES GUIONE S MEDIOS COMO VIÑETAS DENTRO DE ALS CELDAS"
        prompt = prompt & "Se detecto mediante en red privada en sitio red privada, pero explica los escenarios que consideres necesarios" & vbCrLf & Chr(10)
        prompt = prompt & "SOLO DOS columnas, nombre y amenaza"
        prompt = prompt & " " & vbCrLf & Chr(10) & listaVulnerabilidades & vbCrLf & Chr(10)
          
        ' Copiar el prompt generado al portapapeles
        CopiarAlPortapapeles prompt
        
        ' Mostrar un mensaje informativo
        MsgBox "El prompt ha sido copiado al portapapeles.", vbInformation, "Prompt Generado"
    Else
        MsgBox "No se encontraron valores en la selecci?n.", vbExclamation, "Sin Datos"
    End If
End Sub


Sub CYB076_PrepararPromptDesdeSeleccion_PropuestaRemediacionVuln_EnRedPrivada()
    Dim celda As Range
    Dim listaVulnerabilidades As String
    Dim prompt As String
    
    ' Inicializar la variable que almacenar? las vulnerabilidades
    listaVulnerabilidades = ""
    
    ' Recorrer las celdas seleccionadas y acumular los valores con saltos de línea y comas
    For Each celda In Selection
        If Not IsEmpty(celda.value) Then
            If listaVulnerabilidades <> "" Then
                listaVulnerabilidades = listaVulnerabilidades & "," & vbCrLf
            End If
            listaVulnerabilidades = listaVulnerabilidades & celda.value
        End If
    Next celda
    
    ' Construir el prompt final
    If listaVulnerabilidades <> "" Then
        ' prompt = prompt & ""
        prompt = "Hola, por favor Redacta como un pentester un p?rrafo t?cnico de propuesta de remediaci?n que comience con la frase: -Se recomienda...-. Incluye tantos detalles puntuales par remaicion por jemplo nombre de soluciones que funeriona como poroteccon o controles de sguridad, disoitivos, practicas, menciona de manera m?s puntual que se pod´ria hacer para que le encargado del sistema, activo, pueda saber como remediari La respueste debe tener SOLO parrafo breve introducci?n y luego viñetas para lso putnos de propueta de remedaicion Responde para la siguientel lista de vulnerabilidades en FORMATO TABLA DOS COLUMNAS, , SIEMPRE COMIENZAX CON -Se recomienda...- TEXTO AMPLIDO MAS DE 80 palabras, atalmente aplicable a muchos casos, lengauje so framworks"
           prompt = prompt & "Se detecto mediante en red privada en sitio red privada, pero explica los escenarios que consideres necesarios" & vbCrLf & Chr(10)
          prompt = prompt & "Menciona solo soluciones corporativas"
          prompt = prompt & "Solo dos columnas, nombre y propuesat remediacion"
        prompt = prompt & " " & vbCrLf & Chr(10) & listaVulnerabilidades & vbCrLf & Chr(10)
          
        ' Copiar el prompt generado al portapapeles
        CopiarAlPortapapeles prompt
        
        ' Mostrar un mensaje informativo
        MsgBox "El prompt ha sido copiado al portapapeles.", vbInformation, "Prompt Generado"
    Else
        MsgBox "No se encontraron valores en la selecci?n.", vbExclamation, "Sin Datos"
    End If
End Sub

Sub CYB0761_PrepararPromptDesdeSeleccion_ExplicacionTecnicaVuln_EnRedPrivada()
    Dim celda As Range
    Dim listaVulnerabilidades As String
    Dim prompt As String
    
    ' Inicializar la variable que almacenar? las vulnerabilidades
    listaVulnerabilidades = ""
    
    ' Recorrer las celdas seleccionadas y acumular los valores con saltos de línea y comas
    For Each celda In Selection
        If Not IsEmpty(celda.value) Then
            If listaVulnerabilidades <> "" Then
                listaVulnerabilidades = listaVulnerabilidades & "," & vbCrLf
            End If
            listaVulnerabilidades = listaVulnerabilidades & celda.value
        End If
    Next celda
    
    ' Construir el prompt final
    If listaVulnerabilidades <> "" Then
       ' PROMPT EXPLICACIÓN TÉCNICA:
prompt = "Hola, por favor en una tabla, solo dos columnas: vulnerabilidad y explicaci?n t?cnica. "
prompt = prompt & "Para cada una de estas vulnerabilidades redacta un p?rrafo de explicaci?n t?cnica que contenga un ejemplo y "
prompt = prompt & "una conclusi?n breve, concisa y convincente desde la perspectiva de pentesting. "
prompt = prompt & "Inicia la explicaci?n con el texto -En un escenario...- [típico / común / poco probable]. "
prompt = prompt & "De ser posible, agrega c?digo de ejemplo para comprender este tipo de vulnerabilidad. "
prompt = prompt & "El c?digo debe ser útil, no seas escaso en detalles. "
prompt = prompt & "NO MENCIONES RECOMENDACIONES. "
prompt = prompt & "Ejemplo: "
prompt = prompt & "-En un escenario... "
prompt = prompt & "Ejemplo: "
prompt = prompt & "Set-Cookie: sessionID=12345; "
prompt = prompt & "String filePath = -/data/- + userInput + -.txt-; "
prompt = prompt & "public fun{} "
prompt = prompt & "Etc.... "
prompt = prompt & "Se considera inseguro o una vulnerabilidad debido a que... "
prompt = prompt & "En conclusi?n, esta vulnerabilidad es [POTENCIALMENTE EXPLOTABLE] en lo que respecta al c?digo est?tico. "
prompt = prompt & "QUIERO UNA TABLA CON BUEN FORMATO EN LAS CELDAS, SALTOS DE LÍNEA APROPIADOS. "
prompt = prompt & "MÁS DE 125 CARACTERES. "
prompt = prompt & "NO PONGAS TODO EN UN SOLO PÁRRAFO, USA SALTOS DE LÍNEA DENTRO DE LAS CELDAS DE EXPLICACIÓN PARA QUE SEA LEGIBLE. "
        prompt = prompt & " " & vbCrLf & Chr(10) & listaVulnerabilidades & vbCrLf & Chr(10)
        
        ' Copiar el prompt generado al portapapeles
        CopiarAlPortapapeles prompt
        
        ' Mostrar un mensaje informativo
        MsgBox "El prompt ha sido copiado al portapapeles.", vbInformation, "Prompt Generado"
    Else
        MsgBox "No se encontraron valores en la selecci?n.", vbExclamation, "Sin Datos"
    End If
End Sub

Sub CYB0762_PrepararPromptDesdeSeleccion_VectorCVSSVuln_EnRedPrivada()
    Dim celda As Range
    Dim listaVulnerabilidades As String
    Dim prompt As String
    
    ' Inicializar la variable que almacenar? las vulnerabilidades
    listaVulnerabilidades = ""
    
    ' Recorrer las celdas seleccionadas y acumular los valores con saltos de línea y comas
    For Each celda In Selection
        If Not IsEmpty(celda.value) Then
            If listaVulnerabilidades <> "" Then
                listaVulnerabilidades = listaVulnerabilidades & "," & vbCrLf
            End If
            listaVulnerabilidades = listaVulnerabilidades & celda.value
        End If
    Next celda
    
    ' Construir el prompt final
    If listaVulnerabilidades <> "" Then
        prompt = prompt & "Hola, por favor en una tabla, solo dos columnas: vulnerabilidad Severidad. " & vbCrLf
        prompt = prompt & "Exploitability Metrics" & vbCrLf
        prompt = prompt & "Attack Vector (AV): este CAMPO de acuerdo con el tipo de sistema ANALISIS DE VULNERABILIDADES DE INFRAESTRUCTURA DESDE RED PRIVADA" & vbCrLf
        prompt = prompt & "Attack Complexity (AC): " & vbCrLf
        prompt = prompt & "Attack Requirements (AT): " & vbCrLf
        prompt = prompt & "Privileges Required (PR): " & vbCrLf
        prompt = prompt & "User Interaction (UI): " & vbCrLf
        prompt = prompt & "Vulnerable System Impact Metrics" & vbCrLf
        prompt = prompt & "Confidentiality (VC): " & vbCrLf
        prompt = prompt & "Integrity (VI): " & vbCrLf
        prompt = prompt & "Availability (VA): " & vbCrLf
        prompt = prompt & "Subsequent System Impact Metrics" & vbCrLf
        prompt = prompt & "Confidentiality (SC): " & vbCrLf
        prompt = prompt & "Integrity (SI): " & vbCrLf
        prompt = prompt & "Availability (SA): " & vbCrLf
        prompt = prompt & "Exploitation Metrics" & vbCrLf
        prompt = prompt & "Exploitability: " & vbCrLf
        prompt = prompt & "Complexity: " & vbCrLf
        prompt = prompt & "Vulnerable system: " & vbCrLf
        prompt = prompt & "Subsequent system: " & vbCrLf
        prompt = prompt & "Exploitation: " & vbCrLf
        prompt = prompt & "Security requirements: " & vbCrLf
        prompt = prompt & "Por favor evalúe la severidad con base en los criterios anteriores, siendo exigente y estricto al asignar la severidad del CVSS. No asigne impactos altos a menos que haya evidencia de que se pueda explotar directamente y afecte de manera significativa." & vbCrLf
        prompt = prompt & "SOLO RESPONDE CADENAS VECTOR COMPLETAS EJEMPLO CVSS:4.0/AV:A/AC:L/AT:N/PR:N/UI:N/VC:N/VI:N/VA:N/SC:N/SI:N/SA:N VULNEBILIDAD, CVSS" & vbCrLf
        prompt = prompt & "VULNEBILIDAD, CVSS" & vbCrLf
        prompt = prompt & " " & vbCrLf & Chr(10) & listaVulnerabilidades & vbCrLf & Chr(10)
        
        ' Copiar el prompt generado al portapapeles
        CopiarAlPortapapeles prompt
        
        ' Mostrar un mensaje informativo
        MsgBox "El prompt ha sido copiado al portapapeles.", vbInformation, "Prompt Generado"
    Else
        MsgBox "No se encontraron valores en la selecci?n.", vbExclamation, "Sin Datos"
    End If
End Sub

Sub CYB077_PrepararPromptDesdeSeleccion_DescripcionVuln_DesdeInternet()
    Dim celda As Range
    Dim listaVulnerabilidades As String
    Dim prompt As String
    
    ' Inicializar la variable que almacenar? las vulnerabilidades
    listaVulnerabilidades = ""
    
    ' Recorrer las celdas seleccionadas y acumular los valores con saltos de línea y comas
    For Each celda In Selection
        If Not IsEmpty(celda.value) Then
            If listaVulnerabilidades <> "" Then
                listaVulnerabilidades = listaVulnerabilidades & "," & vbCrLf
            End If
            listaVulnerabilidades = listaVulnerabilidades & celda.value
        End If
    Next celda
    
    ' Construir el prompt final
    If listaVulnerabilidades <> "" Then
        prompt = "Redacta un p?rrafo t?cnico breve y conciso que describa la vulnerabilidad detectada, comenzando con la frase: -El sistema...-. Explica en qu? consiste la debilidad de seguridad de manera t?cnica. No incluyas escenarios de explotaci?n, ya que eso corresponde a otro campo. No describas c?mo se explota, solo en qu? consiste el problema. No menciones el nombre exacto de la vulnerabilidad; utiliza expresiones similares. Vulnerabilidades a describi en formato tabla: "
        prompt = prompt & " " & vbCrLf & Chr(10) & listaVulnerabilidades & vbCrLf & Chr(10)
        prompt = prompt & "Contexto de an?lisis: An?lisis de vulnerabilidades de infraestructura a partir conexion a red en desde internet en sitio. Vulnerabilidad detectada mediante escaneo e interacciones desde desde internet en desde internet en sitio."

        
        ' Copiar el prompt generado al portapapeles
        CopiarAlPortapapeles prompt
        
        ' Mostrar un mensaje informativo
        MsgBox "El prompt ha sido copiado al portapapeles.", vbInformation, "Prompt Generado"
    Else
        MsgBox "No se encontraron valores en la selecci?n.", vbExclamation, "Sin Datos"
    End If
End Sub



Sub CYB078_PrepararPromptDesdeSeleccion_AmenazaVuln_DesdeInternet()
    Dim celda As Range
    Dim listaVulnerabilidades As String
    Dim prompt As String
    
    ' Inicializar la variable que almacenar? las vulnerabilidades
    listaVulnerabilidades = ""
    
    ' Recorrer las celdas seleccionadas y acumular los valores con saltos de línea y comas
    For Each celda In Selection
        If Not IsEmpty(celda.value) Then
            If listaVulnerabilidades <> "" Then
                listaVulnerabilidades = listaVulnerabilidades & "," & vbCrLf
            End If
            listaVulnerabilidades = listaVulnerabilidades & celda.value
        End If
    Next celda
    
    ' Construir el prompt final
    If listaVulnerabilidades <> "" Then
        ' prompt = prompt & ""
        prompt = "Hola, por favor considera el siguiente ejemplo: Un atacante podría (obtener, realizar, ejecutar, visualizar, identificar, listar)... Algunos posibles vectores de ataque adicionales asociados con esta amenaza incluyen (pueden ser algunos de los siguientes, pero no necesariamente todos): •  Malware: Un malware diseñado para automatizar intentos de fuerza bruta podría explotar la vulnerabilidad para... •    Usuario malintencionado: Un usuario dentro de la red local con conocimiento de la vulnerabilidad podría aprovecharla para... •    Personal interno: Un empleado con acceso y conocimientos t?cnicos podría, intencionalmente o por error,... •  Delincuente cibern?tico: Un atacante externo en busca de vulnerabilidades podría intentar explotar esta debilidad para... Instrucciones adicionales: 1.   Pregunta si el sistema es interno o externo para determinar los vectores de ataque m?s relevantes, ya que no todos aplican en todos los casos. "
        prompt = prompt & "2.   Redacta una descripci?n de la amenaza que incluya posibles escenarios de ataque, comenzando con la frase: -Un atacante podría...-. 3.    No es necesario proporcionar un ejemplo para cada vector de ataque. Selecciona el m?s realista o probable. 4.   En las viñetas de vectores de ataque, menciona la probabilidad de ocurrencia, incluso si es baja. 5.    El contexto es un an?lisis de vulnerabilidades de infraestructura desde una desde internet, específicamente: -Vulnerabilidad detectada mediante escaneo e interacciones desde desde internet-. Formato de respuesta: •    Responde en una tabla de dos columnas. •    Para cada vulnerabilidad, redacta un p?rrafo descriptivo en la primera columna (mínimo 75 palabras). •  En la segunda columna, lista los vectores de ataque con viñetas (usando guiones  - ). • No uses HTML, solo texto plano. Ejemplo de estructura: Descripci?n de la amenaza    Vectores de ataque Un atacante podría explotar esta vulnerabilidad para acceder _
 informaci?n"
        prompt = prompt & "confidencial Esta amenaza es particularmente crítica en sistemas internos donde los controles de seguridad son menos estrictos."
        prompt = prompt & "Un escenario probable incluye...    - Malware: Un malware podría ser utilizado para... (probabilidad media). - Usuario malintencionado: Un empleado con acceso podría... (probabilidad baja). - Delincuente cibern?tico: Un atacante externo podría... (probabilidad alta). ES MUY IMPORANTE QU EPAR ALOS VECTORI DE ATQUE DE LA MENAZA USSES GUIONE S MEDIOS COMO VIÑETAS DENTRO DE ALS CELDAS"
        prompt = prompt & "Se detecto mediante en desde internet en sitio desde internet, pero explica los escenarios que consideres necesarios" & vbCrLf & Chr(10)
        prompt = prompt & "SOLO DOS columnas, nombre y amenaza"
        prompt = prompt & " " & vbCrLf & Chr(10) & listaVulnerabilidades & vbCrLf & Chr(10)
          
        ' Copiar el prompt generado al portapapeles
        CopiarAlPortapapeles prompt
        
        ' Mostrar un mensaje informativo
        MsgBox "El prompt ha sido copiado al portapapeles.", vbInformation, "Prompt Generado"
    Else
        MsgBox "No se encontraron valores en la selecci?n.", vbExclamation, "Sin Datos"
    End If
End Sub


Sub CYB079_PrepararPromptDesdeSeleccion_PropuestaRemediacionVuln_DesdeInternet()
    Dim celda As Range
    Dim listaVulnerabilidades As String
    Dim prompt As String
    
    ' Inicializar la variable que almacenar? las vulnerabilidades
    listaVulnerabilidades = ""
    
    ' Recorrer las celdas seleccionadas y acumular los valores con saltos de línea y comas
    For Each celda In Selection
        If Not IsEmpty(celda.value) Then
            If listaVulnerabilidades <> "" Then
                listaVulnerabilidades = listaVulnerabilidades & "," & vbCrLf
            End If
            listaVulnerabilidades = listaVulnerabilidades & celda.value
        End If
    Next celda
    
    ' Construir el prompt final
    If listaVulnerabilidades <> "" Then
        ' prompt = prompt & ""
        prompt = "Hola, por favor Redacta como un pentester un p?rrafo t?cnico de propuesta de remediaci?n que comience con la frase: -Se recomienda...-. Incluye tantos detalles puntuales par remaicion por jemplo nombre de soluciones que funeriona como poroteccon o controles de sguridad, disoitivos, practicas, menciona de manera m?s puntual que se pod´ria hacer para que le encargado del sistema, activo, pueda saber como remediari La respueste debe tener SOLO parrafo breve introducci?n y luego viñetas para lso putnos de propueta de remedaicion Responde para la siguientel lista de vulnerabilidades en FORMATO TABLA DOS COLUMNAS, , SIEMPRE COMIENZAX CON -Se recomienda...- TEXTO AMPLIDO MAS DE 80 palabras, atalmente aplicable a muchos casos, lengauje so framworks"
           prompt = prompt & "Se detecto mediante en desde internet en sitio desde internet, pero explica los escenarios que consideres necesarios" & vbCrLf & Chr(10)
          prompt = prompt & "Menciona solo soluciones corporativas"
          prompt = prompt & "Solo dos columnas, nombre y propuesat remediacion"
        prompt = prompt & " " & vbCrLf & Chr(10) & listaVulnerabilidades & vbCrLf & Chr(10)
          
        ' Copiar el prompt generado al portapapeles
        CopiarAlPortapapeles prompt
        
        ' Mostrar un mensaje informativo
        MsgBox "El prompt ha sido copiado al portapapeles.", vbInformation, "Prompt Generado"
    Else
        MsgBox "No se encontraron valores en la selecci?n.", vbExclamation, "Sin Datos"
    End If
End Sub

Sub CYB080_PreparePromptFromSelection_DescripcionVuln_FromCode()
    Dim celda As Range
    Dim listaVulnerabilidades As String
    Dim prompt As String
    
    ' Inicializar la variable que almacenar? las vulnerabilidades
    listaVulnerabilidades = ""
    
    ' Recorrer las celdas seleccionadas y acumular los valores con saltos de línea y comas
    For Each celda In Selection
        If Not IsEmpty(celda.value) Then
            If listaVulnerabilidades <> "" Then
                listaVulnerabilidades = listaVulnerabilidades & "," & vbCrLf
            End If
            listaVulnerabilidades = listaVulnerabilidades & celda.value
        End If
    Next celda
    
    ' Construir el prompt final
    If listaVulnerabilidades <> "" Then
        prompt = ""
        prompt = prompt & "Redacta un p?rrafo t?cnico breve y conciso que describa la vulnerabilidad detectada en el c?digo fuente, comenzando con la frase: " & vbCrLf
        prompt = prompt & "-El c?digo...-. Explica en qu? consiste la debilidad de seguridad de manera t?cnica. " & vbCrLf
        prompt = prompt & "No incluyas escenarios de explotaci?n, ya que eso corresponde a otro campo. " & vbCrLf
        prompt = prompt & "No describas c?mo se explota, solo en qu? consiste el problema. " & vbCrLf
        prompt = prompt & "No menciones el nombre exacto de la vulnerabilidad; utiliza expresiones similares." & vbCrLf & vbCrLf
        prompt = prompt & "Vulnerabilidades a describir en formato tabla:" & vbCrLf
        prompt = prompt & listaVulnerabilidades & vbCrLf & vbCrLf
        prompt = prompt & "Contexto de an?lisis: " & vbCrLf
        prompt = prompt & "An?lisis de c?digo fuente mediante herramientas de an?lisis est?tico (SAST). " & vbCrLf
        prompt = prompt & "Vulnerabilidad identificada a partir del an?lisis del c?digo sin ejecutar la aplicaci?n." & vbCrLf

        ' Copiar el prompt generado al portapapeles
        CopiarAlPortapapeles prompt
        
        ' Mostrar un mensaje informativo
        MsgBox "El prompt ha sido copiado al portapapeles.", vbInformation, "Prompt Generado"
    Else
        MsgBox "No se encontraron valores en la selecci?n.", vbExclamation, "Sin Datos"
    End If
End Sub

Sub CYB081_PreparePromptFromSelection_AmenazaVuln_FromCode()
    Dim celda As Range
    Dim listaVulnerabilidades As String
    Dim prompt As String
    
    ' Inicializar la variable que almacenar? las vulnerabilidades
    listaVulnerabilidades = ""
    
    ' Recorrer las celdas seleccionadas y acumular los valores con saltos de línea y comas
    For Each celda In Selection
        If Not IsEmpty(celda.value) Then
            If listaVulnerabilidades <> "" Then
                listaVulnerabilidades = listaVulnerabilidades & "," & vbCrLf
            End If
            listaVulnerabilidades = listaVulnerabilidades & celda.value
        End If
    Next celda
    
    ' Construir el prompt final
    If listaVulnerabilidades <> "" Then
        prompt = ""
        prompt = prompt & "Hola, por favor considera el siguiente ejemplo: " & vbCrLf
        prompt = prompt & "Un atacante podría (inyectar, manipular, filtrar, exponer, escalar privilegios)..." & vbCrLf
        prompt = prompt & "Algunos posibles vectores de ataque adicionales asociados con esta amenaza incluyen (pueden ser algunos de los siguientes, pero no necesariamente todos):" & vbCrLf
        prompt = prompt & "•  Inyecci?n de c?digo: Una entrada no validada en el c?digo fuente podría permitir inyecci?n de comandos..." & vbCrLf
        prompt = prompt & "•  Exposici?n de informaci?n sensible: Una mala gesti?n de credenciales en el c?digo podría revelar secretos..." & vbCrLf
        prompt = prompt & "•  Elevaci?n de privilegios: Una funci?n mal diseñada podría permitir a un usuario ejecutar acciones con m?s permisos de los necesarios..." & vbCrLf
        prompt = prompt & "•  Manipulaci?n de datos: Un atacante podría modificar par?metros dentro del c?digo para alterar la l?gica de la aplicaci?n..." & vbCrLf
        prompt = prompt & vbCrLf
        prompt = prompt & "Instrucciones adicionales:" & vbCrLf
        prompt = prompt & "1. Pregunta si el c?digo pertenece a una aplicaci?n interna o externa para determinar los vectores de ataque m?s relevantes." & vbCrLf
        prompt = prompt & "2. Redacta una descripci?n de la amenaza que incluya posibles escenarios de ataque, comenzando con la frase: -Un atacante podría...-." & vbCrLf
        prompt = prompt & "3. No es necesario proporcionar un ejemplo para cada vector de ataque. Selecciona el m?s realista o probable." & vbCrLf
        prompt = prompt & "4. En las viñetas de vectores de ataque, menciona la probabilidad de ocurrencia, incluso si es baja." & vbCrLf
        prompt = prompt & "5. El contexto es un an?lisis de c?digo fuente mediante herramientas de an?lisis est?tico (SAST), sin ejecutar la aplicaci?n." & vbCrLf
        prompt = prompt & vbCrLf
        prompt = prompt & "Formato de respuesta:" & vbCrLf
        prompt = prompt & "• Responde en una tabla de dos columnas." & vbCrLf
        prompt = prompt & "• Para cada vulnerabilidad, redacta un p?rrafo descriptivo en la primera columna (mínimo 75 palabras)." & vbCrLf
        prompt = prompt & "• En la segunda columna, lista los vectores de ataque con viñetas (usando guiones - )." & vbCrLf
        prompt = prompt & "• No uses HTML, solo texto plano." & vbCrLf
        prompt = prompt & vbCrLf
        prompt = prompt & "Ejemplo de estructura:" & vbCrLf
        prompt = prompt & "Descripci?n de la amenaza    Vectores de ataque" & vbCrLf
        prompt = prompt & "Un atacante podría explotar esta vulnerabilidad en el c?digo para ejecutar comandos arbitrarios..." & vbCrLf
        prompt = prompt & "    - Inyecci?n de c?digo: Un usuario malintencionado podría insertar c?digo malicioso... (probabilidad media)." & vbCrLf
        prompt = prompt & "    - Exposici?n de informaci?n: Credenciales en el c?digo podrían filtrarse... (probabilidad alta)." & vbCrLf
        prompt = prompt & vbCrLf
        prompt = prompt & "ES MUY IMPORTANTE QUE PARA LOS VECTORES DE ATAQUE DE LA AMENAZA USES GUIONES MEDIOS COMO VIÑETAS DENTRO DE LAS CELDAS." & vbCrLf
        prompt = prompt & "An?lisis realizado mediante herramientas SAST en c?digo fuente est?tico sin ejecuci?n." & vbCrLf
        prompt = prompt & "SOLO DOS COLUMNAS: NOMBRE Y AMENAZA." & vbCrLf
        prompt = prompt & vbCrLf & listaVulnerabilidades & vbCrLf & vbCrLf
          
        ' Copiar el prompt generado al portapapeles
        CopiarAlPortapapeles prompt
        
        ' Mostrar un mensaje informativo
        MsgBox "El prompt ha sido copiado al portapapeles.", vbInformation, "Prompt Generado"
    Else
        MsgBox "No se encontraron valores en la selecci?n.", vbExclamation, "Sin Datos"
    End If
End Sub


Sub CYB082_PreparePromptFromSelection_PropuestaRemediacionVuln_FromCode()
    Dim celda As Range
    Dim listaVulnerabilidades As String
    Dim prompt As String
    
    ' Inicializar la variable que almacenar? las vulnerabilidades
    listaVulnerabilidades = ""
    
    ' Recorrer las celdas seleccionadas y acumular los valores con saltos de línea y comas
    For Each celda In Selection
        If Not IsEmpty(celda.value) Then
            If listaVulnerabilidades <> "" Then
                listaVulnerabilidades = listaVulnerabilidades & "," & vbCrLf
            End If
            listaVulnerabilidades = listaVulnerabilidades & celda.value
        End If
    Next celda
    
    ' Construir el prompt final
    If listaVulnerabilidades <> "" Then
        prompt = "Hola, por favor redacta como un pentester un p?rrafo t?cnico de propuesta de remediaci?n que comience con la frase: -Se recomienda...-."
        prompt = prompt & " Incluye tantos detalles puntuales como sea posible, mencionando soluciones específicas, controles de seguridad, dispositivos y pr?cticas recomendadas."
        prompt = prompt & " Proporciona informaci?n clara para que el encargado del sistema o activo sepa exactamente c?mo remediarlo."
        prompt = prompt & " La respuesta debe contener un p?rrafo breve de introducci?n seguido de viñetas con los puntos de la propuesta de remediaci?n."
        prompt = prompt & " Formato de respuesta: una tabla de dos columnas."
        prompt = prompt & " Siempre comienza con -Se recomienda...-."
        prompt = prompt & " El texto debe ser amplio, con m?s de 80 palabras, aplicable a múltiples casos y en lenguaje t?cnico adecuado."
        prompt = prompt & " Se detect? mediante an?lisis desde internet en el sitio, pero explica los escenarios relevantes."
        prompt = prompt & " Menciona solo soluciones corporativas."
        prompt = prompt & " Solo dos columnas: nombre y propuesta de remediaci?n."
        prompt = prompt & " " & vbCrLf & Chr(10) & listaVulnerabilidades & vbCrLf & Chr(10)
        
        ' Copiar el prompt generado al portapapeles
        CopiarAlPortapapeles prompt
        
        ' Mostrar un mensaje informativo
        MsgBox "El prompt ha sido copiado al portapapeles.", vbInformation, "Prompt Generado"
    Else
        MsgBox "No se encontraron valores en la selecci?n.", vbExclamation, "Sin Datos"
    End If
End Sub


Sub CYB083_Verificar_VectorCVSS4_0()
    Dim celda As Range
    Dim cvssString As String
    Dim url As String
    
    ' Inicializar la variable para almacenar el CVSS vector
    cvssString = ""
    
    ' Recorrer las celdas seleccionadas y acumular los valores
    For Each celda In Selection
        If Not IsEmpty(celda.value) Then
            If cvssString <> "" Then
                cvssString = cvssString & "/" & celda.value
            Else
                cvssString = celda.value
            End If
        End If
    Next celda
    
    ' Verificar si se encontr? un vector CVSS
    If cvssString <> "" Then
        ' Construir la URL
        url = "https://www.first.org/cvss/calculator/4.0#" & cvssString
        
        ' Copiar la URL al portapapeles
        CopiarAlPortapapeles url
        
        ' Mostrar un mensaje con la URL copiada
        MsgBox "La URL ha sido copiada al portapapeles: " & vbCrLf & url, vbInformation, "URL Generada"
        
        ' Abrir la URL en el navegador
        ThisWorkbook.FollowHyperlink url
    Else
        MsgBox "No se encontraron valores en la selecci?n.", vbExclamation, "Sin Datos"
    End If
End Sub


Sub CopiarAlPortapapeles(text As String)
    ' Copiar texto al portapapeles usando Microsoft Forms
    Dim MSForms_DataObject As Object
    Set MSForms_DataObject = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    MSForms_DataObject.SetText text
    MSForms_DataObject.PutInClipboard
    Set MSForms_DataObject = Nothing
End Sub




Public Sub InsertarTextoMarkdownEnWordConFormato(WordApp As Object, WordDoc As Object, placeholder As String, markdownText As String, basePath As String)
    ' --- Constantes de Word para Late Binding (definidas localmente) ---
    Const wdStory As Long = 6
    Const wdFindContinue As Long = 1
    Const wdReplaceOne As Long = 1 ' Reemplaza solo la primera ocurrencia encontrada
    Const wdReplaceAll As Long = 2 ' Reemplaza todas las ocurrencias
    Const wdCollapseEnd As Long = 0
    Const wdLineStyleSingle As Long = 1
    Const wdLineWidth050pt As Long = 4 ' Ancho línea 0.5 pt
    Const wdColorGray25 As Long = 14277081 ' RGB(217, 217, 217) Aprox
    Const wdColorAutomatic As Long = -16777216
    Const wdAlignParagraphCenter As Long = 1
    Const wdAlignParagraphLeft As Long = 0
    Const wdListApplyToSelection As Long = 0
    Const wdListNumberStyleBullet As Long = 23 ' Viñeta común (puede variar según plantilla)
    Const wdContinueList As Long = 1
    Const wdRestartNumbering As Long = 0
    Const wdListNoNumbering As Long = 0 ' Constante para RemoveNumbers
    Const wdFormatDocument As Long = 0 ' Para ApplyListTemplate ListLevelNumber
    Const wdNumberListNum As Long = 2 ' Tipo de lista numerada (puede variar)
    Const wdBulletListNum As Long = 1 ' Tipo de lista con viñetas (puede variar)
    ' Constantes de borde (WdBorderType)
    Const wdBorderTop As Long = -1
    Const wdBorderLeft As Long = -2
    Const wdBorderBottom As Long = -3
    Const wdBorderRight As Long = -4
    Const wdBorderHorizontal As Long = -5 ' Borde entre líneas dentro de la selección
    Const wdBorderVertical As Long = -6   ' Borde entre columnas dentro de la selección

    ' --- Declaración de variables ---
    Dim sel As Object ' La selección actual en Word
    Dim lines() As String ' Array con las líneas del Markdown
    Dim i As Long ' Contador para el bucle de líneas
    Dim lineText As String ' El texto de la línea actual
    Dim trimmedLine As String ' Línea sin espacios al inicio/final
    Dim inCodeBlock As Boolean ' Estado: ¿estamos dentro de un bloque ``` ?
    Dim fs As Object ' FileSystemObject para manejo de archivos/rutas
    Dim fullImgPath As String ' Ruta completa a la imagen (usada en la sub de imágenes)
    Dim codeBlockStartRange As Object ' Rango donde empieza el bloque de código
    Dim currentListType As String ' "bullet", "number", o "" (ninguna)
    Dim lastLineWasList As Boolean ' Para continuar listas
    Dim hLevel As Integer ' Nivel de encabezado (1-6)
    Dim contentText As String ' Texto limpio (sin marcadores iniciales como #, *, 1.)
    Dim listMarkerPos As Integer ' Posición del marcador de lista numerada (.)
    Dim isPathAbsolute As Boolean ' Flag para ruta de imagen absoluta (usada en la sub de imágenes)
    Dim paraRange As Object ' Para referenciar el párrafo actual

    ' --- Manejador de errores global ---
    On Error GoTo ErrorHandler

    ' *** Validación Rigurosa de Objetos de Entrada ***
    If WordApp Is Nothing Then MsgBox "Error crítico: La variable 'WordApp' no representa una instancia válida de Word.", vbCritical, "Error de Parámetro": Exit Sub
    If WordDoc Is Nothing Then MsgBox "Error crítico: La variable 'WordDoc' no representa un documento de Word válido.", vbCritical, "Error de Parámetro": Exit Sub
    On Error Resume Next ' Temporalmente para probar acceso a propiedades
    Dim testAppName As String: testAppName = WordApp.Name
    If Err.Number <> 0 Then MsgBox "Error crítico: La variable 'WordApp' no parece ser una aplicación Word válida (Error: " & Err.Description & ").", vbCritical, "Error de Aplicación Word": Err.Clear: On Error GoTo ErrorHandler: Exit Sub
    Dim testDocName As String: testDocName = WordDoc.Name
    If Err.Number <> 0 Then MsgBox "Error crítico: La variable 'WordDoc' no parece ser un documento Word válido (Error: " & Err.Description & ").", vbCritical, "Error de Documento Word": Err.Clear: On Error GoTo ErrorHandler: Exit Sub
    On Error GoTo ErrorHandler ' Restaurar manejador global

    ' *** Crear FileSystemObject ***
    On Error Resume Next
    Set fs = CreateObject("Scripting.FileSystemObject")
    If Err.Number <> 0 Or fs Is Nothing Then
        MsgBox "Error crítico: No se pudo crear el 'Scripting.FileSystemObject', necesario para manejar rutas de archivo. Verifique que esté registrado en su sistema.", vbCritical, "Error de Creación FSO"
        Exit Sub
    End If
    On Error GoTo ErrorHandler

    ' *** Activar Documento/App y Obtener Selección ***
    On Error Resume Next ' Puede fallar si Word está ocupado o cerrado inesperadamente
    WordDoc.Activate
    WordApp.Activate
    If Err.Number <> 0 Then
        Debug.Print "Advertencia: No se pudo activar la ventana de Word o el documento. Se continuará, pero podría haber problemas si Word no está visible/activo. Error: " & Err.Description
        Err.Clear
    End If
    On Error GoTo ErrorHandler ' Restaurar

    Set sel = WordApp.Selection
    If sel Is Nothing Then MsgBox "Error crítico: No se pudo obtener el objeto 'Selection' de Word.", vbCritical, "Error de Selección": Exit Sub

    ' *** Asegurar que basePath (si se proporciona) termine con separador de ruta ***
    If Len(Trim(basePath)) > 0 Then
        If Right(basePath, 1) <> fs.GetStandardStream(1).Write(vbNullString) And Right(basePath, 1) <> "/" Then ' fs.PathSeparator no existe, truco para obtener '\'
            Dim pathSep As String
             On Error Resume Next ' Acceder a Application puede fallar si WordApp no es válido
             pathSep = WordApp.PathSeparator
             If Err.Number <> 0 Then pathSep = "\" ' Fallback final si WordApp no es válido
             Err.Clear
             On Error GoTo ErrorHandler
            basePath = basePath & pathSep
             Debug.Print "BasePath con separador añadido manualmente: '" & basePath & "'"
         End If
         Debug.Print "BasePath final: '" & basePath & "'"
    Else
        Debug.Print "No se proporcionó BasePath. Las rutas relativas de imágenes podrían fallar si se procesan posteriormente."
    End If


    ' *** 1. Encontrar y Reemplazar el Placeholder ***
    ' Ir al inicio para asegurar que busca desde el principio
    sel.HomeKey wdStory
    sel.Find.ClearFormatting
    sel.Find.Replacement.ClearFormatting
    With sel.Find
        .text = placeholder
        .Replacement.text = "" ' Borrar el placeholder antes de insertar
        .Forward = True
        .Wrap = wdFindContinue ' Buscar en todo el documento desde el inicio
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceOne ' Ejecutar y reemplazar solo la PRIMERA ocurrencia
    End With

    ' Verificar si se encontró y reemplazó
    If sel.Find.found Then
        Debug.Print "Placeholder '" & placeholder & "' encontrado y reemplazado. Insertando contenido Markdown..."

        ' *** 2. Preparar y Procesar el Texto Markdown línea por línea ***
        markdownText = Replace(markdownText, vbCrLf, vbLf) ' Normalizar saltos de línea a LF (Chr(10))
        markdownText = Replace(markdownText, vbCr, vbLf)   ' Normalizar saltos CR a LF
        lines = Split(markdownText, vbLf)

        ' Inicializar estados
        inCodeBlock = False
        currentListType = ""
        lastLineWasList = False
        Set codeBlockStartRange = Nothing

        ' Establecer el estilo Normal por defecto al inicio de la inserción
        On Error Resume Next ' El estilo "Normal" podría no existir o tener otro nombre
        sel.Range.Style = WordDoc.Styles("Normal")
        If Err.Number <> 0 Then
            Debug.Print "Advertencia: No se pudo aplicar el estilo 'Normal'. Se usará el formato por defecto. Error: " & Err.Description
            Err.Clear
        End If
        On Error GoTo ErrorHandler
        sel.Font.Reset ' Resetear formato de fuente explícitamente
        sel.ParagraphFormat.Reset ' Resetear formato de párrafo explícitamente

        ' *** Bucle principal de procesamiento de líneas ***
        For i = 0 To UBound(lines)
            lineText = lines(i)
            trimmedLine = Trim(lineText)

            ' --- Resetear estado de lista si la línea NO es de lista o está vacía Y no estamos en bloque de código ---
            If Not inCodeBlock Then
                 ' Condición para resetear lista: línea vacía O (NO empieza con marcador de lista Y NO es continuación de lista)
                Dim isBulleted As Boolean: isBulleted = (Left(trimmedLine, 1) = "*" Or Left(trimmedLine, 1) = "-" Or Left(trimmedLine, 1) = "+") And Len(trimmedLine) > 1
                Dim isNumbered As Boolean: isNumbered = False
                If IsNumeric(Left(trimmedLine, 1)) Then
                    listMarkerPos = InStr(trimmedLine, ". ")
                    If listMarkerPos > 1 And listMarkerPos <= Len(Left(trimmedLine, 1)) + 2 Then ' Asegura que sea "N. " o "NN. "
                        isNumbered = True
                    End If
                End If

                If Len(trimmedLine) = 0 Or Not (isBulleted Or isNumbered) Then
                    If currentListType <> "" Then
                        ' Salir del modo lista
                        On Error Resume Next
                        ' Intenta quitar el formato de lista del párrafo actual (puede ser redundante si Word ya lo hizo)
                        sel.Range.ListFormat.RemoveNumbers NumberType:=wdListNoNumbering
                        If Err.Number <> 0 Then Debug.Print "Advertencia leve: No se pudo quitar formato de lista explícitamente. Error: " & Err.Description: Err.Clear
                        ' Restablecer formato normal para el párrafo actual si está vacío
                        If Len(Trim(sel.Paragraphs(1).Range.text)) <= 1 Then
                            sel.ParagraphFormat.Reset
                            sel.Font.Reset
                        End If
                        On Error GoTo ErrorHandler
                        currentListType = ""
                        Debug.Print "Fin de lista detectado (Línea no es de lista o vacía)."
                    End If
                    lastLineWasList = False ' Reiniciar flag
                End If
            End If

            ' --- Manejo de Párrafo Vacío (fuera de bloque de código) ---
            If Len(trimmedLine) = 0 And Not inCodeBlock Then
                ' Insertar párrafo vacío solo si el párrafo actual NO está ya vacío
                ' (La lógica anterior ya reseteó la lista si era necesario)
                Dim currentParaText As String
                On Error Resume Next
                currentParaText = sel.Paragraphs(1).Range.text
                If Err.Number <> 0 Then currentParaText = "Error": Err.Clear  ' Asumir no vacío si hay error
                On Error GoTo ErrorHandler

                If Len(Trim(currentParaText)) > 1 Then ' > 1 para ignorar solo el marcador de párrafo
                   sel.TypeParagraph
                   Debug.Print "Insertando párrafo vacío."
                   ' Asegurarse de que el nuevo párrafo tenga formato Normal
                   On Error Resume Next
                   sel.Range.Style = WordDoc.Styles("Normal")
                   If Err.Number <> 0 Then Debug.Print "Adv: No se pudo aplicar estilo Normal a párrafo vacío.": Err.Clear
                   On Error GoTo ErrorHandler
                   sel.Font.Reset
                   sel.ParagraphFormat.Reset
                Else
                   Debug.Print "Párrafo ya vacío, omitiendo inserción extra."
                   ' Si estamos en una línea vacía pero el párrafo anterior era de lista,
                   ' nos aseguramos de que este párrafo vacío no tenga formato de lista.
                   If lastLineWasList Then
                       On Error Resume Next
                       sel.Range.ListFormat.RemoveNumbers NumberType:=wdListNoNumbering
                       sel.ParagraphFormat.Reset ' Quitar sangría de lista
                       sel.Font.Reset
                       Err.Clear
                       On Error GoTo ErrorHandler
                       currentListType = "" ' Marcar que ya no estamos en lista
                       lastLineWasList = False
                   End If
                End If
                GoTo NextLine ' Procesar siguiente línea
            End If ' Fin manejo línea vacía

            ' --- Manejo de Bloques de Código (```) ---
            If trimmedLine = "```" Then
                inCodeBlock = Not inCodeBlock
                If inCodeBlock Then ' Iniciando bloque
                    ' Insertar un párrafo ANTES del bloque si el actual no está vacío
                    Dim isEmptyParaCodeStart As Boolean
                    On Error Resume Next
                    isEmptyParaCodeStart = (Len(Trim(sel.Paragraphs(1).Range.text)) <= 1)
                    If Err.Number <> 0 Then isEmptyParaCodeStart = False: Err.Clear
                    On Error GoTo ErrorHandler
                    If Not isEmptyParaCodeStart Then sel.TypeParagraph

                    Set codeBlockStartRange = sel.Paragraphs(1).Range ' Marcar inicio del rango (el párrafo actual)
                    ' Aplicar formato al párrafo actual (donde empezará el código)
                    sel.Font.Name = "Courier New"
                    sel.Font.Size = 10
                    sel.Font.Bold = False ' Asegurar no negrita
                    sel.Font.Italic = False ' Asegurar no cursiva
                    sel.ParagraphFormat.SpaceBefore = 6
                    sel.ParagraphFormat.SpaceAfter = 0
                    sel.ParagraphFormat.LeftIndent = WordApp.InchesToPoints(0.25) ' Añadir una pequeña sangría
                    sel.ParagraphFormat.RightIndent = WordApp.InchesToPoints(0.25)
                    Debug.Print "Inicio de bloque de código detectado."
                Else ' Cerrando bloque de código
                    If Not codeBlockStartRange Is Nothing Then
                        Dim endParaRange As Object
                        Dim blockRange As Object
                        On Error Resume Next ' El párrafo actual podría ser el último del doc

                        ' El párrafo actual es el que contiene el ``` de cierre.
                        ' Queremos el rango desde el inicio del párrafo de apertura
                        ' hasta el *final* del párrafo *anterior* al de cierre.
                        Set endParaRange = sel.Paragraphs(1).Previous.Range
                        If Err.Number <> 0 Or endParaRange Is Nothing Then
                           ' Si no hay párrafo anterior (bloque de 1 línea o error),
                           ' el rango es solo el párrafo inicial.
                           Set blockRange = codeBlockStartRange
                           Debug.Print "Advertencia: Bloque de código corto o error al obtener párrafo anterior. Formateando solo el párrafo inicial."
                           Err.Clear
                        Else
                           ' Rango desde el inicio del párrafo de apertura hasta el final del párrafo anterior al cierre
                           Set blockRange = WordDoc.Range(Start:=codeBlockStartRange.Start, End:=endParaRange.End)
                        End If
                        On Error GoTo ErrorHandler

                        If Not blockRange Is Nothing Then
                            ' Aplicar borde y sombreado al rango completo del bloque
                            If blockRange.Start < blockRange.End Or blockRange.Characters.Count > 1 Then ' Asegurarse rango válido
                                On Error Resume Next ' Operaciones de formato pueden fallar
                                With blockRange.ParagraphFormat.Borders
                                     .Enable = True ' Activar bordes para el párrafo(s)
                                     ' Definir borde exterior
                                     Dim borderType As Variant
                                     For Each borderType In Array(wdBorderTop, wdBorderLeft, wdBorderBottom, wdBorderRight)
                                         With .Item(CLng(borderType))
                                             .LineStyle = wdLineStyleSingle
                                             .LineWidth = wdLineWidth050pt
                                             .Color = wdColorGray25
                                         End With
                                     Next borderType
                                     ' Quitar bordes internos (por si acaso)
                                     .Item(wdBorderHorizontal).LineStyle = 0 ' wdLineStyleNone = 0
                                     .Item(wdBorderVertical).LineStyle = 0   ' wdLineStyleNone = 0
                                End With
                                blockRange.Shading.BackgroundPatternColor = RGB(248, 248, 248) ' Gris muy claro
                                If Err.Number <> 0 Then Debug.Print "Adv: Error parcial al aplicar formato al bloque de código. " & Err.Description: Err.Clear
                                Debug.Print "Bloque de código formateado. Rango: " & blockRange.Start & "-" & blockRange.End
                            Else
                                Debug.Print "Advertencia: Rango de bloque de código inválido o vacío (" & blockRange.Start & "-" & blockRange.End & "). No se aplicó formato de borde/sombreado."
                            End If
                            Err.Clear ' Limpiar errores de formato
                            On Error GoTo ErrorHandler
                            Set blockRange = Nothing ' Liberar
                        End If

                        ' Mover cursor después del bloque y resetear formato
                        sel.Collapse wdCollapseEnd ' Mover al final del párrafo del ```
                        sel.TypeParagraph ' Nuevo párrafo para contenido normal
                        sel.ParagraphFormat.Reset ' Resetear formato párrafo
                        sel.Font.Reset          ' Resetear formato fuente
                         ' Asegurarse de quitar borde y sombreado del nuevo párrafo explícitamente
                        On Error Resume Next
                        sel.ParagraphFormat.Borders.Enable = False
                        sel.Shading.BackgroundPatternColor = wdColorAutomatic
                        Err.Clear
                        On Error GoTo ErrorHandler
                        Debug.Print "Fin de bloque de código. Formato reseteado."
                    Else
                        Debug.Print "Advertencia: Se encontró '```' de cierre sin uno de apertura registrado."
                        ' Insertar un párrafo para evitar que el texto siguiente se pegue
                        sel.TypeParagraph
                    End If
                    Set codeBlockStartRange = Nothing ' Resetear marcador de inicio
                End If
                lastLineWasList = False ' Un bloque de código rompe la lista

            ElseIf inCodeBlock Then
                ' Dentro de un bloque de código: Insertar texto tal cual
                ' No usar ProcessInlineFormatting aquí
                sel.TypeText text:=lineText
                sel.TypeParagraph ' Siguiente línea dentro del bloque
                ' Asegurar que la fuente/formato se mantenga (Word a veces la resetea)
                sel.Font.Name = "Courier New"
                sel.Font.Size = 10
                sel.Font.Bold = False
                sel.Font.Italic = False
                sel.ParagraphFormat.LeftIndent = WordApp.InchesToPoints(0.25) ' Mantener sangría
                sel.ParagraphFormat.RightIndent = WordApp.InchesToPoints(0.25)
                sel.ParagraphFormat.SpaceBefore = 0 ' Sin espacio extra entre líneas de código
                sel.ParagraphFormat.SpaceAfter = 0
                lastLineWasList = False

            ' --- Manejo de Encabezados (# a ######) ---
            ElseIf Left(trimmedLine, 1) = "#" Then
                hLevel = 0
                Do While Left(trimmedLine, hLevel + 1) Like String(hLevel + 1, "#") And hLevel < 6
                    hLevel = hLevel + 1
                Loop

                ' Verificar si hay un espacio después de los #
                If hLevel > 0 And Mid(trimmedLine, hLevel + 1, 1) = " " Then
                    contentText = Trim(Mid(trimmedLine, hLevel + 2)) ' Texto después de "# "
                    Debug.Print "Encabezado Nivel " & hLevel & " detectado: '" & contentText & "'"

                    ' Asegurarse de estar en un párrafo nuevo si el actual no está vacío
                    Set paraRange = sel.Paragraphs(1).Range
                    If Len(Trim(paraRange.text)) > 1 Then sel.TypeParagraph

                    ' *** Aplicar SOLO negrita ***
                    sel.Font.Bold = True
                    ' Opcional: Aplicar tamaños decrecientes (pero no estilos)
                    Select Case hLevel
                        Case 1: sel.Font.Size = 16
                        Case 2: sel.Font.Size = 14
                        Case 3: sel.Font.Size = 13
                        Case 4: sel.Font.Size = 12
                        Case Else: sel.Font.Size = 11
                    End Select
                    sel.ParagraphFormat.SpaceBefore = IIf(hLevel <= 2, 12, 6) ' Espacio antes de headers
                    sel.ParagraphFormat.SpaceAfter = IIf(hLevel <= 3, 6, 4)  ' Espacio después de headers

                    ' Insertar el texto del encabezado procesando formato inline (que puede estar dentro del header)
                    ' ProcessInlineFormatting NO añadirá TypeParagraph si isHeader = True
                    ProcessInlineFormatting sel, contentText, isHeader:=True

                    ' Mover al final de la línea insertada y añadir párrafo
                    sel.Collapse wdCollapseEnd
                    sel.TypeParagraph

                    ' *** Resetear formato para el siguiente párrafo ***
                    sel.Font.Reset ' Quita negrita, tamaño, etc.
                    sel.ParagraphFormat.Reset ' Quita espacio antes/después, etc.
                    On Error Resume Next ' Aplicar estilo Normal si existe
                    sel.Range.Style = WordDoc.Styles("Normal")
                    Err.Clear
                    On Error GoTo ErrorHandler

                    lastLineWasList = False ' Encabezado rompe lista
                Else ' No parece un encabezado válido (ej. #sin espacio)
                    Debug.Print "Tratando línea que empieza con # pero no es encabezado como texto normal."
                     Set paraRange = sel.Paragraphs(1).Range ' Comprobar si párrafo está vacío
                     If Len(Trim(paraRange.text)) > 1 Then sel.TypeParagraph ' Si no está vacío, empezar nuevo párrafo
                     sel.ParagraphFormat.Reset ' Resetear formato para texto normal
                     sel.Font.Reset
                     ProcessInlineFormatting sel, trimmedLine
                     lastLineWasList = False
                End If

            ' --- Manejo de Listas con Viñetas (*, -, +) ---
            ElseIf (Left(trimmedLine, 1) = "*" Or Left(trimmedLine, 1) = "-" Or Left(trimmedLine, 1) = "+") And Mid(trimmedLine, 2, 1) = " " Then
                contentText = Trim(Mid(trimmedLine, 3)) ' Quitar el marcador y espacio
                Debug.Print "Elemento de lista con viñeta detectado: '" & contentText & "'"

                Set paraRange = sel.Paragraphs(1).Range
                If Len(Trim(paraRange.text)) > 1 And Not lastLineWasList Then
                    sel.TypeParagraph ' Nuevo párrafo si el actual no está vacío y no es continuación de lista
                End If

                ' Aplicar/Continuar formato de lista
                If currentListType <> "bullet" Or Not lastLineWasList Then ' Nueva lista o cambio de tipo
                    On Error Resume Next
                    ' Usar ApplyListTemplate para más control si ApplyBulletDefault falla
                    Dim listGalleryBullet As Object
                    Set listGalleryBullet = WordApp.ListGalleries(wdBulletListNum) ' wdBulletGallery = 2
                    sel.Range.ListFormat.ApplyListTemplateWithLevel ListTemplate:=listGalleryBullet.ListTemplates(1), ContinuePreviousList:=False, ApplyTo:=wdListApplyToSelection
                    If Err.Number <> 0 Then
                       Debug.Print "Error aplicando plantilla de viñeta (" & Err.Description & "). Intentando ApplyBulletDefault."
                       Err.Clear
                       sel.Range.ListFormat.ApplyBulletDefault
                       If Err.Number <> 0 Then
                            Debug.Print "Error aplicando formato de viñeta. Insertando texto con tabulación."
                            Err.Clear
                            sel.TypeText vbTab ' Simular indentación
                        Else
                           currentListType = "bullet"
                           Debug.Print "Aplicado formato de lista con viñetas (ApplyBulletDefault)."
                       End If
                    Else
                        currentListType = "bullet"
                        Debug.Print "Aplicado formato de lista con viñetas (ApplyListTemplate)."
                    End If
                    On Error GoTo ErrorHandler
                Else ' Continuación de lista
                    ' Word debería manejarlo automáticamente al escribir, pero forzar párrafo si no se creó antes
                    If sel.Start = paraRange.Start And Len(Trim(paraRange.text)) <= 1 Then
                         ' No hacer nada, Word insertará la viñeta al escribir
                    ElseIf sel.Start > paraRange.Start Or Len(Trim(paraRange.text)) > 1 Then
                         sel.TypeParagraph ' Forzar nuevo item si no estamos al inicio de un párrafo vacío
                    End If
                     Debug.Print "Continuando lista con viñetas."
                End If

                ' Insertar el texto del elemento procesando formato inline
                ' isListItem=True evitará que ProcessInlineFormatting añada TypeParagraph extra
                ProcessInlineFormatting sel, contentText, isListItem:=True
                lastLineWasList = True

            ' --- Manejo de Listas Numeradas (1., 2., etc.) ---
            ElseIf isNumbered Then ' Variable calculada al inicio del bucle
                contentText = Trim(Mid(trimmedLine, listMarkerPos + 1)) ' Texto después de "N. "
                Debug.Print "Elemento de lista numerada detectado: '" & contentText & "'"

                 Set paraRange = sel.Paragraphs(1).Range
                 If Len(Trim(paraRange.text)) > 1 And Not lastLineWasList Then
                     sel.TypeParagraph ' Nuevo párrafo si el actual no está vacío y no es continuación de lista
                 End If

                 ' Aplicar/Continuar formato de lista numerada
                 If currentListType <> "number" Or Not lastLineWasList Then ' Nueva lista o cambio de tipo
                    On Error Resume Next
                     ' Usar ApplyListTemplate para más control
                     Dim listGalleryNumber As Object
                     Set listGalleryNumber = WordApp.ListGalleries(wdNumberListNum) ' wdNumberGallery = 3
                     sel.Range.ListFormat.ApplyListTemplateWithLevel ListTemplate:=listGalleryNumber.ListTemplates(1), ContinuePreviousList:=False, ApplyTo:=wdListApplyToSelection, DefaultListBehavior:=wdWord10ListBehavior ' Ajustar comportamiento
                    If Err.Number <> 0 Then
                         Debug.Print "Error aplicando plantilla numerada (" & Err.Description & "). Intentando ApplyNumberDefault."
                         Err.Clear
                         sel.Range.ListFormat.ApplyNumberDefault
                         If Err.Number <> 0 Then
                            Debug.Print "Error aplicando formato numerado. Insertando texto con número y tab."
                            Err.Clear
                            sel.TypeText Trim(Left(trimmedLine, listMarkerPos)) & vbTab ' Insertar número y tab
                         Else
                            currentListType = "number"
                            Debug.Print "Aplicado formato de lista numerada (ApplyNumberDefault)."
                        End If
                    Else
                       currentListType = "number"
                       Debug.Print "Aplicado formato de lista numerada (ApplyListTemplate)."
                    End If
                    On Error GoTo ErrorHandler
                 Else ' Continuación de lista
                    ' Word debería manejarlo, forzar párrafo si es necesario
                    If sel.Start = paraRange.Start And Len(Trim(paraRange.text)) <= 1 Then
                         ' No hacer nada
                    ElseIf sel.Start > paraRange.Start Or Len(Trim(paraRange.text)) > 1 Then
                        sel.TypeParagraph ' Forzar nuevo item
                    End If
                    Debug.Print "Continuando lista numerada."
                 End If

                ' Insertar el texto del elemento procesando formato inline
                ProcessInlineFormatting sel, contentText, isListItem:=True
                lastLineWasList = True

            ' --- Texto Normal (Párrafo) ---
            Else
                Debug.Print "Procesando como texto normal: '" & trimmedLine & "'"
                Set paraRange = sel.Paragraphs(1).Range
                ' Si el párrafo actual no está vacío Y (la línea anterior no era lista O estamos al inicio de un texto nuevo), empezar nuevo párrafo.
                 If Len(Trim(paraRange.text)) > 1 And (Not lastLineWasList Or i = 0) Then
                    sel.TypeParagraph
                    sel.ParagraphFormat.Reset ' Asegurar formato normal
                    sel.Font.Reset
                 End If
                 ' Asegurar formato normal si venimos de una lista
                 If lastLineWasList Then
                    sel.ParagraphFormat.Reset
                    sel.Font.Reset
                 End If

                ProcessInlineFormatting sel, trimmedLine
                lastLineWasList = False ' Texto normal rompe lista (ya se hizo al inicio del loop, pero redundancia segura)
                currentListType = ""    ' Asegurar que no hay lista activa
            End If

NextLine:       ' Etiqueta para saltar al final del bucle si era línea vacía
        Next i
        ' ------ Fin del Bucle Principal ------

        ' --- Limpieza Final después del bucle ---
        ' Si el último elemento fue un bloque de código sin cerrar
        If inCodeBlock Then
            Debug.Print "Advertencia: El texto Markdown terminó dentro de un bloque de código sin cierre '```'."
            ' Intentar aplicar formato final al bloque incompleto
             If Not codeBlockStartRange Is Nothing Then
                 Dim lastRange As Object
                 Set lastRange = WordDoc.Range(Start:=codeBlockStartRange.Start, End:=sel.Range.End)
                 On Error Resume Next
                 With lastRange.ParagraphFormat.Borders
                      .Enable = True
                      Dim borderTypeEnd As Variant
                      For Each borderTypeEnd In Array(wdBorderTop, wdBorderLeft, wdBorderBottom, wdBorderRight)
                          With .Item(CLng(borderTypeEnd))
                             .LineStyle = wdLineStyleSingle
                             .LineWidth = wdLineWidth050pt
                             .Color = wdColorGray25
                          End With
                      Next borderTypeEnd
                      .Item(wdBorderHorizontal).LineStyle = 0
                      .Item(wdBorderVertical).LineStyle = 0
                 End With
                 lastRange.Shading.BackgroundPatternColor = RGB(248, 248, 248)
                 If Err.Number <> 0 Then Debug.Print "Adv: Error formato bloque final: " & Err.Description: Err.Clear
                 On Error GoTo ErrorHandler
                 Set lastRange = Nothing
            End If
            ' Resetear formato para cualquier cosa que venga después
            sel.Collapse wdCollapseEnd
            sel.TypeParagraph ' Asegurar un párrafo final limpio
            sel.Font.Reset
            sel.ParagraphFormat.Reset
            On Error Resume Next
            sel.ParagraphFormat.Borders.Enable = False
            sel.Shading.BackgroundPatternColor = wdColorAutomatic
            Err.Clear
            On Error GoTo ErrorHandler
        End If

        ' Si el último elemento fue parte de una lista, quitar el formato del último párrafo (que suele estar vacío)
        If currentListType <> "" Then
            On Error Resume Next
             Set paraRange = sel.Paragraphs(1).Range
             If Len(Trim(paraRange.text)) <= 1 Then ' Solo si el último párrafo está vacío
                 paraRange.ListFormat.RemoveNumbers NumberType:=wdListNoNumbering
                 paraRange.ParagraphFormat.Reset ' Quitar sangría
                 If Err.Number <> 0 Then Debug.Print "Advertencia: No se pudo quitar formato de lista final explícitamente. Error: " & Err.Description: Err.Clear
             End If
            On Error GoTo ErrorHandler
        End If

        Debug.Print "Procesamiento de Markdown completado."

    Else
        MsgBox "Error: No se encontró el placeholder '" & placeholder & "' en el documento.", vbExclamation, "Placeholder No Encontrado"
    End If ' Fin de If sel.Find.Found

    ' Liberar objetos
    Set sel = Nothing
    Set fs = Nothing
    Set codeBlockStartRange = Nothing
    Set paraRange = Nothing

    Exit Sub ' Salida normal

ErrorHandler:
    ' Mostrar un mensaje de error más detallado
    Dim errMsg As String
    errMsg = "Se produjo un error inesperado en 'InsertarMarkdownEnWord'." & vbCrLf & vbCrLf & _
             "Error Número: " & Err.Number & vbCrLf & _
             "Descripción: " & Err.Description & vbCrLf & _
             "Fuente: " & Err.Source & vbCrLf & vbCrLf
    ' Intentar añadir información de línea si i está en un rango válido
    On Error Resume Next ' Evitar error si i no está inicializado o fuera de rango
    errMsg = errMsg & "Última línea procesada (aprox.): " & i
    If i >= 0 And i <= UBound(lines) Then
       errMsg = errMsg & " -> '" & lines(i) & "'"
    End If
    If Err.Number <> 0 Then
        errMsg = errMsg & " (No se pudo determinar el contenido de la línea)."
        Err.Clear
    End If
    On Error GoTo 0 ' Desactivar manejo de errores temporalmente

    MsgBox errMsg, vbCritical, "Error en Ejecución VBA"

    ' Limpiar objetos por si acaso antes de salir
    On Error Resume Next ' Ignorar errores durante la limpieza
    Set sel = Nothing
    Set fs = Nothing
    Set codeBlockStartRange = Nothing
    Set paraRange = Nothing
    On Error GoTo 0 ' Desactivar manejo de errores antes de salir

End Sub



Private Sub ProcessInlineFormatting(sel As Object, text As String, Optional isListItem As Boolean = False, Optional isHeader As Boolean = False)
    ' --- Constantes locales ---
    Const BOLD_MARKER As String = "**"
    Const ITALIC_MARKER As String = "*" ' Importante: Usar '*' para cursiva, no '_'
    Const CODE_MARKER As String = "`"
    Const ESCAPE_CHAR As String = "\" ' Caracter de escape
    Const CODE_FONT_NAME As String = "Courier New"
    Const CODE_FONT_SIZE As Long = 10
    Const WD_COLLAPSE_END As Long = 0 ' Definir localmente
    Dim wdColorAutomaticInline As Long: wdColorAutomaticInline = -16777216 ' Definir localmente

    ' --- Variables ---
    Dim currPos As Long
    Dim char As String, nextChar As String, prevChar As String
    Dim textToInsert As String
    Dim initialFont As Object ' Para recordar fuente/tamaño base del párrafo
    Dim initialBold As Boolean, initialItalic As Boolean ' Estado inicial antes del código inline
    Dim isCurrentlyCode As Boolean ' Estado local para fuente de código
    Dim tempRange As Object ' Para obtener la fuente base

    On Error GoTo InlineErrorHandler

    ' --- 1. Preparar Selección y Guardar Estado Inicial de Fuente ---
    ' No añadir párrafo aquí, se maneja en el bucle principal.
    ' Solo guardar el estado de la fuente actual ANTES de procesar inline.
    Set tempRange = sel.Range ' Clonar selección actual para no moverla
    Set initialFont = tempRange.Font.Duplicate ' Guardar una copia del formato de fuente
    Set tempRange = Nothing
    isCurrentlyCode = False ' Asegurar que no empezamos en modo código

    ' --- 2. Bucle de Procesamiento Inline ---
    currPos = 1
    textToInsert = ""

    Do While currPos <= Len(text)
        char = Mid(text, currPos, 1)
        ' Mirar el siguiente y anterior carácter
        If currPos < Len(text) Then nextChar = Mid(text, currPos + 1, 1) Else nextChar = ""
        If currPos > 1 Then prevChar = Mid(text, currPos - 1, 1) Else prevChar = ""

        ' --- Manejo de Caracter de Escape (\) ---
        If char = ESCAPE_CHAR Then
             ' Si el siguiente es un marcador o \ , añadir el siguiente y saltar ambos
            If nextChar = Left(BOLD_MARKER, 1) Or nextChar = ITALIC_MARKER Or nextChar = CODE_MARKER Or nextChar = ESCAPE_CHAR Then
                textToInsert = textToInsert & nextChar
                currPos = currPos + 2 ' Saltar \ y el carácter escapado
                GoTo ContinueLoopInline
            Else ' Escape de otro carácter (no especial), añadir el escape y el carácter
                textToInsert = textToInsert & char ' Añadir la barra invertida también
                ' El carácter siguiente se añadirá en la siguiente iteración como normal
                currPos = currPos + 1
                GoTo ContinueLoopInline
            End If
        End If

        ' --- Detectar Marcadores (priorizar **) ---
        Dim markerFound As Boolean: markerFound = False

        ' Negrita (**)
        If char = ITALIC_MARKER And nextChar = ITALIC_MARKER Then
            If Len(textToInsert) > 0 Then sel.TypeText textToInsert: textToInsert = "" ' Insertar texto acumulado
            sel.Font.Bold = Not sel.Font.Bold ' Alternar negrita
            currPos = currPos + 2 ' Saltar **
            markerFound = True
        ' Cursiva (*) - Asegurarse que NO sea parte de **
        ElseIf char = ITALIC_MARKER Then
            ' Si el anterior NO era *, entonces este es un * de inicio/fin de cursiva
            If prevChar <> ITALIC_MARKER Then
                 If Len(textToInsert) > 0 Then sel.TypeText textToInsert: textToInsert = ""
                 sel.Font.Italic = Not sel.Font.Italic ' Alternar cursiva
                 currPos = currPos + 1 ' Saltar *
                 markerFound = True
            Else ' El anterior ERA *, este es el segundo * de **, ya procesado. Solo saltar.
                 currPos = currPos + 1
                 markerFound = True ' Evita que se añada al texto normal
            End If
        ' Código Inline (`)
        ElseIf char = CODE_MARKER Then
            If Len(textToInsert) > 0 Then sel.TypeText textToInsert: textToInsert = ""
            isCurrentlyCode = Not isCurrentlyCode ' Alternar estado
            If isCurrentlyCode Then
                ' Guardar estado actual antes de aplicar formato código
                initialBold = sel.Font.Bold
                initialItalic = sel.Font.Italic
                ' Aplicar formato código
                sel.Font.Name = CODE_FONT_NAME
                sel.Font.Size = CODE_FONT_SIZE
                sel.Font.Bold = False
                sel.Font.Italic = False
                 ' Opcional: Añadir un sombreado ligero al código inline
                 sel.Shading.BackgroundPatternColor = RGB(240, 240, 240)
            Else ' Salir de modo código
                ' Resetear a la fuente/tamaño base guardada al inicio de la función
                 If Not initialFont Is Nothing Then
                     sel.Font.Name = initialFont.Name
                     sel.Font.Size = initialFont.Size
                     sel.Shading.BackgroundPatternColor = wdColorAutomaticInline ' Quitar sombreado
                     ' Restaurar negrita/cursiva que había ANTES del código
                     sel.Font.Bold = initialBold
                     sel.Font.Italic = initialItalic
                 Else ' Fallback si falló al guardar initialFont
                     sel.Font.Reset
                     sel.Shading.BackgroundPatternColor = wdColorAutomaticInline
                 End If
            End If
            currPos = currPos + 1 ' Saltar `
            markerFound = True
        End If

        ' --- Acumular texto normal ---
        If Not markerFound Then
            textToInsert = textToInsert & char
            currPos = currPos + 1
        End If

ContinueLoopInline:
    Loop ' Fin del Do While

    ' Insertar cualquier texto restante acumulado
    If Len(textToInsert) > 0 Then sel.TypeText textToInsert

    ' --- 3. Finalizar la línea (NO añadir TypeParagraph si es item de lista o header) ---
    If Not isListItem And Not isHeader Then
        ' Solo para texto normal que no sea parte de lista/header
        sel.Collapse WD_COLLAPSE_END
        sel.TypeParagraph ' Añadir párrafo al final de líneas de texto normales
    End If
    ' Para items de lista (isListItem=True) y headers (isHeader=True),
    ' el TypeParagraph se maneja en el bucle principal DESPUÉS de llamar a esta función.

    Set initialFont = Nothing ' Liberar objeto
    Exit Sub

InlineErrorHandler:
    MsgBox "Error en 'ProcessInlineFormatting':" & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.Description & vbCrLf & _
           "Procesando texto (inicio): '" & Left(text, 50) & "...'", vbCritical, "Error de Formato Inline"
    ' Intentar insertar el texto sin formato como fallback
    On Error Resume Next
    sel.TypeText text ' Insertar el texto original sin formato
    If Not isListItem And Not isHeader Then sel.TypeParagraph ' Añadir párrafo si era texto normal
    Err.Clear
    On Error GoTo 0 ' Desactivar manejo de errores local
    Set initialFont = Nothing
End Sub





Sub FusionarDocumentosInsertando(WordApp As Object, documentsList As Variant, finalDocumentPath As String)
    Dim baseDoc     As Object
    Dim sFile       As String
    Dim oRng        As Object
    Dim i           As Integer
    
    On Error GoTo err_Handler
    
    ' Crear un nuevo documento base
    Set baseDoc = WordApp.Documents.Add
    
    ' Iterar sobre la lista de documentos a fusionar
    For i = LBound(documentsList) To UBound(documentsList)
        sFile = documentsList(i)
        
        ' Insertar el contenido del documento actual al final del documento base
        Set oRng = baseDoc.Range
        oRng.Collapse 0        ' Colapsar el rango al final del documento base
        oRng.InsertFile sFile, , , , True ' Mantener formato original
        
        ' Insertar un salto de página después de cada documento insertado (excepto el último)
        If i < UBound(documentsList) Then
            Set oRng = baseDoc.Range
            oRng.Collapse 0        ' Colapsar el rango al final del documento base
            oRng.InsertBreak Type:=6 ' Insertar un salto de página
        End If
    Next i
    
    ' Guardar el archivo final
    baseDoc.SaveAs finalDocumentPath
    
    ' Cerrar el documento base
    baseDoc.Close
    
    ' Limpiar objetos
    Set baseDoc = Nothing
    Set oRng = Nothing
    
    Exit Sub
    
err_Handler:
    MsgBox "Error: " & Err.Number & vbCrLf & Err.Description
    Err.Clear
    Exit Sub
End Sub







Private Function IsAbsolutePath(path As String, fs As Object) As Boolean
    On Error Resume Next
    IsAbsolutePath = False
    
    ' Validar si fs es válido
    If fs Is Nothing Then
        Debug.Print "FSO inválido en IsAbsolutePath"
        Exit Function
    End If

    ' Limpiar y convertir el path a minúsculas
    Dim lowerPath As String: lowerPath = LCase(Trim(path))

    ' Verificar si es un URL absoluto
    If Left(lowerPath, 7) = "http://" Or Left(lowerPath, 8) = "https://" Then
        IsAbsolutePath = True
        GoTo CleanExitAbsPath
    End If

    ' Verificar si es una ruta UNC
    Dim driveName As String
    driveName = fs.GetDriveName(path)

    ' Verificar si es UNC o una ruta local absoluta
    If Err.Number = 0 Then
        ' Verificar si es una ruta UNC
        If Left(driveName, 2) = "\\" Then
            IsAbsolutePath = True
        ' Verificar si es una ruta local absoluta (por ejemplo, "F:")
        ElseIf Len(driveName) = 2 And Right(driveName, 1) = ":" Then
            If Asc(LCase(Left(driveName, 1))) >= 97 And Asc(LCase(Left(driveName, 1))) <= 122 Then
                IsAbsolutePath = True
            End If
        End If
    End If
    Err.Clear

CleanExitAbsPath:
    ' Limpiar cualquier error no manejado
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Function


Sub RawPrint(inputText As String)
    Dim saltoCarro As Integer
    Dim saltoLinea As Integer
    Dim formattedText As String
    
    ' Inicializar contadores
    saltoCarro = 0
    saltoLinea = 0
    formattedText = inputText
    
    ' Contar saltos de carro y saltos de línea
    saltoCarro = Len(inputText) - Len(Replace(inputText, vbCr, ""))
    saltoLinea = Len(inputText) - Len(Replace(inputText, vbLf, ""))
    
    ' Reemplazar saltos de línea y salto de carro con símbolos visibles
    formattedText = Replace(formattedText, vbCrLf, "[CRLF]")
    formattedText = Replace(formattedText, vbCr, "[CR]")
    formattedText = Replace(formattedText, vbLf, "[LF]")
    
    ' Imprimir el texto con los saltos de línea visibles
    Debug.Print "Texto formateado con saltos visibles:" & vbCrLf & formattedText
    Debug.Print vbCrLf & "Resumen de saltos:"
    Debug.Print "Cantidad de saltos de carro (CR): " & saltoCarro
    Debug.Print "Cantidad de saltos de línea (LF): " & saltoLinea
End Sub






Public Sub SustituirTextoMarkdownPorImagenes(WordApp As Object, WordDoc As Object, basePath As String)
    Dim docRange As Object          ' Word.Range completo del documento
    Dim searchRange As Object       ' Rango para buscar el marcador (se ajusta en cada iteración)
    Dim tagRange As Object          ' Rango que contiene solo el tag Markdown ![...]()
    Dim expandedTagRange As Object  ' Rango expandido para incluir saltos de línea/párrafo adyacentes
    Dim fs As Object                ' FileSystemObject
    Dim imgAltText As String        ' Texto alternativo que se usará como caption
    Dim imgPath As String           ' Ruta extraída (cruda) del tag Markdown
    Dim fullImgPath As String       ' Ruta completa, ya procesada
    Dim altEndMarkerPos As Long     ' Posición donde finaliza el ALT, antes de ]( (relativa a tempSearchRange)
    Dim pathEndMarkerPos As Long    ' Posición de final de la ruta, es decir, del carácter ')' (relativa a tempSearchRange)
    Dim inlineShape As Object       ' Objeto para la imagen insertada (InlineShape)
    Dim originalTagText As String   ' El texto original del tag Markdown (posible con LFs/CRs)
    Dim fileExists As Boolean       ' Bandera para comprobar la existencia del archivo
    Dim continueSearching As Boolean
    Dim pathSep As String           ' Separador de carpeta (de Word)
    Dim charCodeBefore As Long      ' Código del carácter antes del tag
    Dim charCodeAfter As Long       ' Código del carácter después del tag
    Dim currentSearchStart As Long  ' Posición desde donde iniciar la búsqueda en cada iteración
    Dim tempSearchRange As Object   ' Rango temporal para buscar dentro del tag
    Dim coreTagText As String       ' Texto del tag ![alt](path)
    Dim relAltEndPos As Long, relPathStartPos As Long
    Dim insertionPointRange As Object ' Punto donde insertar la imagen
    Dim rngAfterPic As Object       ' Rango después de la imagen insertada
    Dim addPictureErrNum As Long    ' Para capturar error específico de AddPicture
    Dim pathIsAbsolute As Boolean   ' Para chequeo de ruta

    ' Constantes para los códigos de caracteres especiales en Word
    Const CR As Long = 13 ' Código ASCII para Carriage Return (Salto de párrafo ¶)
    Const LF As Long = 10 ' Código ASCII para Line Feed (Salto de línea manual - Shift+Enter)

    ' Constantes para los marcadores Markdown
    Const START_MARKER As String = "!["
    Const ALT_END_MARKER As String = "]("
    Const PATH_END_MARKER As String = ")"

    ' Constantes Word (Late Binding) - Definidas localmente para independencia
    Const wdFindStop As Long = 0
    Const wdCollapseStart As Long = 1
    Const wdCollapseEnd As Long = 0
    Const wdAlignParagraphCenter As Long = 1
    Const wdParagraph As Long = 4 ' Para MoveStart

    ' --- Manejador de Errores Principal ---
    On Error GoTo GlobalErrorHandler

    ' --- Validación Inicial ---
    If WordApp Is Nothing Or WordDoc Is Nothing Then
        MsgBox "Error: La aplicación Word o el Documento no son válidos.", vbCritical, "Error de Entrada"
        Exit Sub
    End If

    ' --- Crear FileSystemObject ---
    On Error Resume Next ' Intentar crear FSO
    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs Is Nothing Or Err.Number <> 0 Then
        On Error GoTo 0 ' Desactivar Resume Next
        MsgBox "Error crítico: No se pudo crear el FileSystemObject.", vbCritical, "Error FSO"
        Exit Sub
    End If
    On Error GoTo GlobalErrorHandler ' Restaurar manejador de errores normal

    ' --- Obtener separador de ruta y normalizar basePath ---
    pathSep = WordApp.PathSeparator
    basePath = Trim(basePath)
    If Len(basePath) > 0 Then
        ' Quitar separadores al final si existen
        Do While Right(basePath, 1) = "\" Or Right(basePath, 1) = "/"
            If Len(basePath) = 1 Then
                basePath = "" ' Evitar bucle infinito si es solo "\" o "/"
                Exit Do
            End If
            basePath = Left(basePath, Len(basePath) - 1)
        Loop
        ' Añadir separador al final si basePath no quedó vacío
        If Len(basePath) > 0 Then
            basePath = basePath & pathSep
        End If
        Debug.Print "BasePath normalizado: """ & basePath & """"
    End If

    ' --- Inicializar búsqueda ---
    currentSearchStart = 0 ' Empezar desde el inicio del documento
    continueSearching = True
    WordApp.ScreenUpdating = False ' Desactivar actualización para rapidez

    ' --- Bucle Principal de Búsqueda y Reemplazo ---
    Do While continueSearching
        ' Define el rango para esta búsqueda específica, desde la última posición
        Set searchRange = WordDoc.Range(Start:=currentSearchStart, End:=WordDoc.content.End)

        searchRange.Find.ClearFormatting
        searchRange.Find.text = START_MARKER
        searchRange.Find.Forward = True
        searchRange.Find.Wrap = wdFindStop ' Detener al final del documento
        searchRange.Find.MatchCase = False ' Ignorar mayúsculas/minúsculas

        If searchRange.Find.Execute Then ' Encontró el inicio del tag "!["
            Dim foundRange As Object
            Set foundRange = searchRange.Duplicate ' foundRange AHORA es donde se encontró "!["

            ' Preparar la próxima búsqueda para DESPUÉS de este hallazgo inicial
            ' Se actualizará más adelante si se procesa el tag, si no, buscará desde aquí
            currentSearchStart = foundRange.End

            ' Buscar los marcadores de fin "](" y ")" DENTRO del texto que sigue a "!["
            Set tempSearchRange = WordDoc.Range(Start:=foundRange.End, End:=WordDoc.content.End)
            ' Usar vbTextCompare para ignorar mayúsculas/minúsculas en marcadores si se desea, aunque no es estándar MD
            altEndMarkerPos = InStr(1, tempSearchRange.text, ALT_END_MARKER, vbTextCompare)

            If altEndMarkerPos > 0 Then ' Encontró "]("
                 ' Buscar ")" DESPUÉS de "]("
                pathEndMarkerPos = InStr(altEndMarkerPos + Len(ALT_END_MARKER), tempSearchRange.text, PATH_END_MARKER, vbTextCompare)

                If pathEndMarkerPos > 0 Then ' Encontró ")" también, tenemos un tag completo
                    ' --- Definir el rango exacto del tag ![alt](path) ---
                    Dim tagStartDocPos As Long, tagEndDocPos As Long
                    tagStartDocPos = foundRange.Start ' Inicio del "!["
                    ' End del tag es el inicio del rango de búsqueda temporal + pos final - 1 (por base 1 de InStr) + longitud del marcador final
                    tagEndDocPos = tempSearchRange.Start + pathEndMarkerPos - 1 + Len(PATH_END_MARKER)

                    Set tagRange = WordDoc.Range(Start:=tagStartDocPos, End:=tagEndDocPos)
                    coreTagText = tagRange.text ' Contiene ![alt](path) potencialmente con saltos de línea internos

                    ' --- Extraer ALT y Path del coreTagText ---
                    relAltEndPos = InStr(1, coreTagText, ALT_END_MARKER, vbTextCompare)
                    If relAltEndPos = 0 Then
                        Debug.Print "Error interno: No se encontró ALT_END_MARKER en coreTagText. Saltando."
                        ' currentSearchStart ya está en foundRange.End, buscará el próximo "!["
                        GoTo ContinueNextFind ' Saltar al final del bucle Do While
                    End If

                    relPathStartPos = relAltEndPos + Len(ALT_END_MARKER)
                    imgAltText = Trim(Mid(coreTagText, Len(START_MARKER) + 1, relAltEndPos - Len(START_MARKER) - 1))
                    imgPath = Trim(Mid(coreTagText, relPathStartPos, pathEndMarkerPos - relPathStartPos)) 'pathEndMarkerPos es relativo a tempSearchRange, no a coreTagText. Recalcular.
                    ' Corrección para extraer imgPath correctamente desde coreTagText
                    Dim relPathEndPos As Long
                    relPathEndPos = InStr(relPathStartPos, coreTagText, PATH_END_MARKER, vbTextCompare)
                    If relPathEndPos > 0 Then
                       imgPath = Trim(Mid(coreTagText, relPathStartPos, relPathEndPos - relPathStartPos))
                    Else
                       Debug.Print "Error interno: No se encontró PATH_END_MARKER en coreTagText. Saltando."
                       GoTo ContinueNextFind ' Saltar al final del bucle Do While
                    End If

                    Debug.Print "Tag encontrado: """ & Replace(Replace(coreTagText, Chr(CR), "[CR]"), Chr(LF), "[LF]") & """"
                    Debug.Print "Texto Alt extraído: """ & imgAltText & """"
                    Debug.Print "Ruta extraída (cruda): """ & imgPath & """"

                    If Len(imgPath) = 0 Then
                        Debug.Print "Advertencia: Ruta de imagen vacía. Saltando tag."
                        currentSearchStart = tagRange.End ' Avanzar después del tag vacío encontrado
                        GoTo ContinueNextFind
                    End If

                    ' --- Expandir rango para incluir CR (saltos de párrafo) adyacentes si existen ---
                    ' Esto ayuda a eliminar el párrafo donde estaba el tag si estaba solo
                    Set expandedTagRange = tagRange.Duplicate
                    originalTagText = expandedTagRange.text ' Guardar antes de expandir
                    Debug.Print "Rango tag inicial: " & expandedTagRange.Start & "-" & expandedTagRange.End

                    ' Comprobar carácter ANTES
                    If expandedTagRange.Start > 0 Then ' Asegurar que no estamos al inicio absoluto
                        On Error Resume Next ' Por si Start=1
                         Dim rngBefore As Object
                         Set rngBefore = WordDoc.Range(expandedTagRange.Start - 1, expandedTagRange.Start)
                         If Err.Number = 0 Then
                            charCodeBefore = AscW(rngBefore.Characters(1).text) ' Usar AscW para Unicode
                            If charCodeBefore = CR Then
                                expandedTagRange.Start = expandedTagRange.Start - 1
                                Debug.Print "CR encontrado antes. Rango expandido a: " & expandedTagRange.Start
                            End If
                         End If
                         Set rngBefore = Nothing
                         Err.Clear
                        On Error GoTo GlobalErrorHandler ' Restaurar
                    End If

                    ' Comprobar carácter DESPUÉS
                    If expandedTagRange.End < WordDoc.content.End Then ' Asegurar que no estamos al final absoluto
                         On Error Resume Next ' Por si End es el último carácter
                         Dim rngAfter As Object
                         Set rngAfter = WordDoc.Range(expandedTagRange.End, expandedTagRange.End + 1)
                          If Err.Number = 0 Then
                            charCodeAfter = AscW(rngAfter.Characters(1).text) ' Usar AscW para Unicode
                            If charCodeAfter = CR Then
                                expandedTagRange.End = expandedTagRange.End + 1
                                Debug.Print "CR encontrado después. Rango expandido a: " & expandedTagRange.End
                            End If
                          End If
                          Set rngAfter = Nothing
                          Err.Clear
                         On Error GoTo GlobalErrorHandler ' Restaurar
                    End If
                    Debug.Print "Texto a reemplazar (expandido): """ & Replace(Replace(expandedTagRange.text, Chr(CR), "[CR]"), Chr(LF), "[LF]") & """"

                    ' --- Construir ruta completa y verificar existencia ---
                    fullImgPath = ""
                    fileExists = False
                    ' Reemplazar barras inclinadas por el separador del sistema
                    imgPath = Replace(imgPath, "/", pathSep)
                    imgPath = Replace(imgPath, "\", pathSep) ' Asegurar consistencia

                    ' Verificar si la ruta es absoluta (local o UNC) o URL
                    On Error Resume Next ' Intentar con FSO.IsAbsolutePath
                    pathIsAbsolute = fs.IsAbsolutePath(imgPath)
                    If Err.Number <> 0 Then ' Fallback manual si FSO falla o no está disponible
                        Err.Clear
                        Debug.Print "Advertencia: fs.IsAbsolutePath falló. Usando chequeo manual."
                        pathIsAbsolute = (InStr(imgPath, ":\") > 0 Or Left(imgPath, 2) = "\\")
                    End If
                    On Error GoTo GlobalErrorHandler ' Restaurar

                    If pathIsAbsolute Then
                        fullImgPath = imgPath
                        On Error Resume Next ' Chequear existencia con FSO
                        fileExists = fs.fileExists(fullImgPath)
                        If Err.Number <> 0 Then
                            Debug.Print "Error FSO chequeando existencia de: " & fullImgPath
                            fileExists = False ' Asumir que no existe si FSO falla
                            Err.Clear
                        End If
                        On Error GoTo GlobalErrorHandler
                    ElseIf Left(LCase(imgPath), 4) = "http" Then ' Es una URL
                        fullImgPath = imgPath
                        fileExists = True ' Asumir OK, AddPicture lo verifica realmente
                        Debug.Print "Ruta es URL: """ & fullImgPath & """"
                    ElseIf Len(basePath) > 0 Then ' Es relativa y tenemos basePath
                        On Error Resume Next
                        fullImgPath = fs.BuildPath(basePath, imgPath)
                        If Err.Number = 0 Then
                           fileExists = fs.fileExists(fullImgPath)
                           If Err.Number <> 0 Then
                                Debug.Print "Error FSO chequeando existencia de ruta construida: " & fullImgPath
                                fileExists = False
                                Err.Clear
                           End If
                        Else ' Error en BuildPath
                           Debug.Print "Error FSO en BuildPath con: """ & basePath & """ y """ & imgPath & """"
                           fullImgPath = basePath & imgPath ' Intento manual simple
                           fileExists = False ' No se pudo construir bien, probablemente no exista
                           Err.Clear
                        End If
                        On Error GoTo GlobalErrorHandler
                        Debug.Print "Ruta relativa construida: """ & fullImgPath & """"
                    Else ' Es relativa pero NO tenemos basePath
                        Debug.Print "Advertencia: Ruta relativa (""" & imgPath & """) encontrada pero no se proporcionó BasePath."
                        fullImgPath = imgPath ' Guardar la ruta relativa por si acaso
                        fileExists = False
                    End If

                    ' --- Insertar Imagen o Mantener Texto ---
                    If fileExists And Len(fullImgPath) > 0 Then
                        Debug.Print "Archivo/URL parece existir: """ & fullImgPath & """. Intentando insertar..."
                        Set insertionPointRange = expandedTagRange.Duplicate
                        insertionPointRange.Collapse Direction:=wdCollapseStart ' Colapsar al inicio del rango a reemplazar

                        expandedTagRange.text = "" ' Borrar texto original (tag + CRs opcionales)

                        ' *** Intento de inserción con manejo de error específico ***
                        Set inlineShape = Nothing
                        addPictureErrNum = 0
                        On Error Resume Next ' Activar manejo de error SOLO para AddPicture
                        Set inlineShape = insertionPointRange.InlineShapes.AddPicture( _
                            fileName:=fullImgPath, LinkToFile:=False, SaveWithDocument:=True)
                        addPictureErrNum = Err.Number ' Guardar número de error si ocurrió
                        On Error GoTo GlobalErrorHandler ' !!! Restaurar manejador global INMEDIATAMENTE !!!

                        If addPictureErrNum = 0 And Not inlineShape Is Nothing Then
                            ' --- Éxito al insertar ---
                            Debug.Print "Imagen insertada con éxito."
                            inlineShape.AlternativeText = imgAltText ' Establecer texto alternativo en la imagen

                            ' --- Añadir Caption (Pie de foto) ---
                            Set rngAfterPic = inlineShape.Range ' Obtener el rango de la imagen insertada
                            rngAfterPic.Collapse Direction:=wdCollapseEnd ' Mover al final de la imagen
                            
                            ' Insertar párrafo después Y luego escribir en él para evitar problemas de formato
                            rngAfterPic.InsertParagraphAfter
                            ' Mover al nuevo párrafo creado
                            rngAfterPic.MoveStart unit:=wdParagraph, Count:=1
                            rngAfterPic.MoveEnd unit:=wdParagraph, Count:=1 ' Asegurar que el rango cubre solo el nuevo párrafo
                            
                            rngAfterPic.text = imgAltText ' Poner el texto del pie de foto
                            rngAfterPic.ParagraphFormat.Alignment = wdAlignParagraphCenter ' Centrar pie de foto
                            rngAfterPic.Font.Italic = True ' Opcional: Poner cursiva al pie de foto
                            rngAfterPic.Font.Size = 9     ' Opcional: Tamaño más pequeño para pie de foto

                            ' La próxima búsqueda debe empezar DESPUÉS del pie de foto
                            currentSearchStart = rngAfterPic.End

                        Else
                            ' --- Fallo al insertar (Error 4120 u otro) ---
                            Debug.Print "*** ERROR " & addPictureErrNum & " al insertar imagen: """ & fullImgPath & """ ***"
                            Debug.Print "Descripción del error: " & Err.Description ' Mostrar descripción del error específico
                            ' Opcional: Insertar un marcador de error en el documento en lugar de la imagen
                            insertionPointRange.InsertAfter "[ERROR AL INSERTAR IMAGEN: " & imgPath & "]"
                            ' La próxima búsqueda debe empezar DESPUÉS del punto de inserción original
                            ' (donde estaba el tag que falló)
                            currentSearchStart = insertionPointRange.End + Len("[ERROR AL INSERTAR IMAGEN: " & imgPath & "]")
                            Err.Clear ' Limpiar el error para continuar
                        End If
                        ' Liberar objetos de esta iteración de inserción
                        Set insertionPointRange = Nothing
                        Set rngAfterPic = Nothing
                        Set inlineShape = Nothing
                    Else
                        ' --- Archivo NO encontrado o ruta vacía ---
                        Debug.Print "Archivo NO encontrado o ruta inválida: """ & fullImgPath & """. Se mantiene el texto original."
                        ' Dejar el texto original (expandedTagRange NO fue borrado)
                        ' La próxima búsqueda debe empezar DESPUÉS del tag original que no se reemplazó
                        currentSearchStart = expandedTagRange.End
                    End If

                    ' Liberar objetos del tag procesado
                    Set expandedTagRange = Nothing
                    Set tagRange = Nothing

                Else ' No se encontró ")" después de "]("
                    Debug.Print "Tag incompleto: Se encontró '!["; y; "](' pero no el ')' final. Saltando."
                    ' Avanzar la búsqueda después del "](" encontrado
                    currentSearchStart = tempSearchRange.Start + altEndMarkerPos + Len(ALT_END_MARKER) - 1
                End If ' pathEndMarkerPos > 0
            Else ' No se encontró "](" después de "!["
                 ' Tag inválido o texto normal que empieza con "!["
                 ' currentSearchStart ya está en foundRange.End, buscará el próximo "!["
                 Debug.Print "Marcador '![' encontrado pero no seguido por ']('. Tratado como texto normal."
            End If ' altEndMarkerPos > 0

            ' Liberar objetos de búsqueda de esta iteración
             Set tempSearchRange = Nothing
             Set foundRange = Nothing

        Else ' Find.Execute no encontró más "!["
            continueSearching = False ' Salir del bucle Do While
            Debug.Print "No se encontraron más marcadores '!['."
        End If ' Find.Execute

ContinueNextFind: ' Etiqueta para saltar aquí si hay error recuperable o tag inválido
        DoEvents ' Permitir que Word procese eventos (útil en bucles largos)
    Loop ' While continueSearching

    GoTo CleanExit

GlobalErrorHandler:
    ' Mostrar un mensaje de error más detallado para errores NO manejados específicamente
    MsgBox "Se produjo un error inesperado en 'SustituirTextoMarkdownPorImagenes'." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & vbCrLf & _
           "Fuente: " & Err.Source & vbCrLf & vbCrLf & _
           "Última ruta procesada (aprox): """ & fullImgPath & """", vbCritical, "Error Global en Ejecución"
    ' Intentar restaurar ScreenUpdating antes de salir
    If Not WordApp Is Nothing Then
       On Error Resume Next ' Ignorar error si WordApp ya no es válido
       WordApp.ScreenUpdating = True
       On Error GoTo 0
    End If
    GoTo ForcedCleanup ' Ir a la limpieza forzada

CleanExit:
    Debug.Print "--- Fin de SustituirTextoMarkdownPorImagenes (Normal) ---"
    ' Salida normal

ForcedCleanup:
    ' Liberar todos los objetos para evitar memoria colgada
    On Error Resume Next ' Ignorar errores durante la limpieza
    If Not fs Is Nothing Then Set fs = Nothing
    Set docRange = Nothing
    Set searchRange = Nothing
    Set tagRange = Nothing
    Set expandedTagRange = Nothing
    Set inlineShape = Nothing
    Set tempSearchRange = Nothing
    Set insertionPointRange = Nothing
    Set rngAfterPic = Nothing
    ' Restaurar actualización de pantalla
    If Not WordApp Is Nothing Then WordApp.ScreenUpdating = True
    On Error GoTo 0 ' Desactivar cualquier manejo de errores restante
End Sub
