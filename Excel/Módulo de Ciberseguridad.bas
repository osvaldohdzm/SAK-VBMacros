Attribute VB_Name = "ExcelModuloCiberseguridad"
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
                    ' Agregamos la línea con los ":" pero sin guión
                    textoConGuiones = textoConGuiones & lineas(i) & vbLf
                    incluirGuion = True ' Habilitamos la adición de guiones después de encontrar el ":"
                ElseIf incluirGuion Then
                    ' Después del primer ":", agregamos un guion
                    If Len(Trim(lineas(i))) > 0 Then
                        textoConGuiones = textoConGuiones & " - " & lineas(i) & vbLf
                    Else
                        ' Si la línea está vacía, solo agregamos el salto de línea
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
    
    ' Mostrar cuadro de diálogo para seleccionar la carpeta de destino
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Selecciona la carpeta de salida"
        If .Show = -1 Then
            carpetaSalida = .SelectedItems(1)
        Else
            MsgBox "No se seleccionó ninguna carpeta.", vbExclamation
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
            
            ' Verificar si se encontró la columna "Severidad"
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
                MsgBox "No se encontró la columna        'Severidad'.", vbExclamation
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
            
            ' Verificar si se encontró la columna "Severidad"
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
                MsgBox "No se encontró la columna        'Severidad'.", vbExclamation
            End If
        End With
        
        wb.Sheets(1).Name = "Vulnerabilidades"
        wb.Sheets(1).Range("A1").Select
        ' Guardar el archivo en la carpeta seleccionada
        wb.SaveAs tempFileName, xlOpenXMLWorkbook
        wb.Close False
        
        FunExportarHojaActivaAExcelINAI = True
        MsgBox "La hoja ha sido exportada con éxito a " & tempFileName, vbInformation
    Else
        FunExportarHojaActivaAExcelINAI = False
        MsgBox "No hay ninguna hoja activa para exportar.", vbExclamation
    End If
    
    Exit Function
    
ErrorHandler:
    FunExportarHojaActivaAExcelINAI = False
    MsgBox "Ocurrió un error: " & Err.Description, vbCritical
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
            ' Convierte el contenido en un array separado por el carácter de nueva línea
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
            
            ' Convertir la colección de claves en un array
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
            
            ' Convierte el array nuevamente en una cadena concatenada por el carácter de nueva línea
            content = Join(uniqueArray, Chr(10))
            
            ' Asigna el contenido filtrado a la celda
            cell.value = content
        End If
    Next cell
End Sub

Sub CYB029_ReplaceWithURLs()
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

Sub CYB030_AplicarFormatoCondicional()
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
    Dim rng         As Range
    Dim celda       As Range
    Dim lineas      As Variant
    Dim i           As Integer
    Dim htmlPattern As String
    Dim liPattern   As String
    Dim cleanHtmlPattern As String
    Dim NuevoTexto As String
    Dim Texto As String
    Dim cleanOutput As String
    Dim lastLineWasEmpty As Boolean

    ' Definir los patrones HTML
    htmlPattern = "<(\/?(p|a|ul|b|strong|i|u|br)[^>]*?)>|<\/p><p>"
    liPattern = "<li[^>]*?>"
    cleanHtmlPattern = "<[^>]+>"        ' Para eliminar cualquier etiqueta HTML y sus atributos
    
    ' Asume que el rango seleccionado es el que quieres modificar
    Set rng = Selection
    
    ' Reemplaza caracteres de tabulación con espacios y quita espacios en blanco adicionales
    For Each celda In rng
        If Not IsEmpty(celda.value) Then
            If Not celda.HasFormula Then ' Ignora celdas con fórmulas
            
                ' Reemplazar caracteres de tabulación con espacios
                celda.value = Replace(celda.value, Chr(9), " ")
                ' Quitar espacios en blanco adicionales
                celda.value = Application.Trim(celda.value)
                
                ' Reemplazar <li> etiquetas con saltos de línea
                celda.value = RegExpReplace(celda.value, liPattern, vbLf)
                
                ' Eliminar otras etiquetas HTML pero conservar el texto interno
                celda.value = RegExpReplace(celda.value, cleanHtmlPattern, vbNullString)
                
                ' Unificar saltos de línea y eliminar líneas vacías consecutivas
                lineas = Split(celda.value, vbLf)
                cleanOutput = ""
                lastLineWasEmpty = False
                
                For i = LBound(lineas) To UBound(lineas)
                    ' Si la línea no está vacía, se agrega directamente
                    If Trim(lineas(i)) <> "" Then
                        cleanOutput = cleanOutput & lineas(i) & vbLf
                        lastLineWasEmpty = False
                    ElseIf Not lastLineWasEmpty Then
                        ' Si la línea está vacía pero no es consecutiva, agregar un solo salto de línea
                        cleanOutput = cleanOutput & vbLf
                        lastLineWasEmpty = True
                    End If
                Next i
                
                ' Eliminar el último salto de línea si existe
                If Len(cleanOutput) > 0 And Right(cleanOutput, 1) = vbLf Then
                    cleanOutput = Left(cleanOutput, Len(cleanOutput) - 1)
                End If
                
                ' Reemplazar diferentes tipos de saltos de línea y retornos de carro
                Texto = ReplaceHtmlEntities(cleanOutput)
                
                ' Mantener los saltos de línea correctos y evitar duplicados
                NuevoTexto = Replace(Texto, vbCrLf, "")   ' Salto de línea + retorno de carro
                NuevoTexto = Replace(NuevoTexto, vbCr, "") ' Retorno de carro
                NuevoTexto = Replace(Texto, Chr(10) & Chr(10), Chr(10))

                
                ' Asignar el texto limpio a la celda
                celda.value = NuevoTexto ' Mantener el formato con saltos de línea intactos
            End If
        End If
    Next celda
    
    MsgBox "Proceso completado: Espacios, etiquetas HTML y saltos de línea ajustados.", vbInformation
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


Sub CYB011_ONLY_URLS_HTTP()

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
            
            ' Verificar si el array no está vacío antes de redimensionar
            If UBound(lineas) >= 0 Then
                ' Crear un nuevo array para almacenar las líneas no vacías
                ReDim lineasSinVacias(0 To UBound(lineas))
                idx = 0
                
                ' Iterar sobre cada línea del array
                For i = LBound(lineas) To UBound(lineas)
                    ' Verificar si la línea está vacía y no agregarla al nuevo array
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
        
        ' Verificar que la celda no esté vacía
        If ip <> "" Then
            respuesta = False ' Inicializar como no respondida
            
            ' Intentar ping hasta 3 veces
            For i = 1 To 3
                ' Ejecutar el ping y capturar la salida
                Set objExec = objShell.Exec("ping -n 1 -w 500 " & ip)
                resultado = objExec.StdOut.ReadAll
                
                ' Si encuentra "TTL=", la IP respondió
                If InStr(1, resultado, "TTL=", vbTextCompare) > 0 Then
                    respuesta = True
                    Exit For ' Salir del bucle si ya respondió
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






Function RegExpReplace(ByVal text As String, ByVal replacePattern As String, ByVal replaceWith As String) As String
    ' Función para reemplazar utilizando expresiones regulares
    Dim regEx       As Object
    Set regEx = CreateObject("VBScript.RegExp")
    
    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .pattern = replacePattern
    End With
    
    RegExpReplace = regEx.Replace(text, replaceWith)
End Function

Function ReplaceHtmlEntities(ByVal text As String) As String
    ' Función para reemplazar entidades HTML con caracteres correspondientes
    text = Replace(text, "&lt;", "<")
    text = Replace(text, "&gt;", ">")
    text = Replace(text, "&amp;", "&")
    text = Replace(text, "&quot;", """")
    text = Replace(text, "&apos;", "'")        ' Comillas simples
    text = Replace(text, "&#x27;", "'")        ' Comillas simples
    text = Replace(text, "&#34;", """")        ' Comillas dobles
    text = Replace(text, "&#39;", "'")         ' Comillas simples
    text = Replace(text, "&#160;", Chr(160))   ' Espacio no separable
    
    ReplaceHtmlEntities = text
End Function


Sub CYB026_OrdenaSegunColorRelleno()
    Dim celdaActual As Range
    Dim tabla       As ListObject
    Dim ws          As Worksheet
    Dim respuesta   As VbMsgBoxResult
    
    ' Obtener la celda actualmente seleccionada
    Set celdaActual = ActiveCell
    
    ' Obtener la hoja activa
    Set ws = ActiveSheet
    
    ' Verificar si la celda seleccionada está dentro de una tabla
    On Error Resume Next
    Set tabla = celdaActual.ListObject
    On Error GoTo 0
    
    ' Confirmar si la tabla encontrada es la tabla de vulnerabilidades
    If Not tabla Is Nothing Then
        ' Mostrar mensaje de confirmación
        respuesta = MsgBox("¿Estás seguro de que estás en una tabla de vulnerabilidades? Procederá a ordenar por color de relleno en la columna        'Severidad'.", vbYesNo + vbQuestion, "Confirmación")
        
        ' Si el usuario elige 'Sí', proceder con la ordenación
        If respuesta = vbYes Then
            With ws.ListObjects(tabla.Name).Sort
                ' Limpiar campos de ordenación previos
                .SortFields.Clear
                
                ' Ordenar por color de relleno verde
                .SortFields.Add key:=ws.ListObjects(tabla.Name).ListColumns("Severidad").Range, _
                SortOn:=xlSortOnCellColor, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal
                .SortFields(1).SortOnValue.Color = RGB(0, 176, 80)
                .Apply
                
                ' Limpiar campos de ordenación previos
                .SortFields.Clear
                
                ' Ordenar por color de relleno amarillo
                .SortFields.Add key:=ws.ListObjects(tabla.Name).ListColumns("Severidad").Range, _
                SortOn:=xlSortOnCellColor, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal
                .SortFields(1).SortOnValue.Color = RGB(255, 255, 0)
                .Apply
                
                ' Limpiar campos de ordenación previos
                .SortFields.Clear
                
                ' Ordenar por color de relleno rojo
                .SortFields.Add key:=ws.ListObjects(tabla.Name).ListColumns("Severidad").Range, _
                SortOn:=xlSortOnCellColor, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal
                .SortFields(1).SortOnValue.Color = RGB(255, 0, 0)
                .Apply
            End With
        End If
    Else
        MsgBox "La celda seleccionada no está dentro de una tabla.", vbExclamation, "Error"
    End If
End Sub
' GenerarDocumentosWord

Sub CYB014_EliminarUltimasFilasSiEsSalidaPruebaSeguridad(WordDoc As Object, replaceDic As Object)
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
    
    ' Verificar si las claves están presentes en el diccionario
    If replaceDic.Exists(salidaPruebaSeguridadKey) And replaceDic.Exists(metodoDeteccionKey) Then
        ' Obtener los valores de las claves
        salidaPruebaSeguridadValue = CStr(replaceDic(salidaPruebaSeguridadKey))
        metodoDeteccionValue = CStr(replaceDic(metodoDeteccionKey))
        
        ' Inicializar la tabla
        Set firstTable = WordDoc.Tables(1)
        numRows = firstTable.Rows.Count
        
        ' Verificar si ambos valores están vacíos
        If Len(Trim(salidaPruebaSeguridadValue)) = 0 And Len(Trim(metodoDeteccionValue)) = 0 Then
            ' Si ambos están vacíos, eliminar las últimas filas de la tabla principal
            If numRows > 0 Then
                ' Eliminar la última fila
                firstTable.Rows(numRows).Delete
                ' Eliminar la penúltima fila si hay más de una fila
                If numRows > 1 Then
                    firstTable.Rows(numRows - 1).Delete
                End If
            End If
      ElseIf Len(Trim(salidaPruebaSeguridadValue)) = 0 Then
            ' Si "Método de detección" tiene texto, eliminar la tabla interna dentro de la última celda
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
    Set nuevoEstilo = docWord.Styles.Add(Name:=estilo, Type:=1)        ' Tipo 1 = Estilo de párrafo
    If Err.Number <> 0 Then
        MsgBox "No se pudo crear el estilo        '" & estilo & "'. " & Err.Description, vbExclamation
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Sub CYB016_ExportarTablaContenidoADocumentoWord()
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
    
    ' Obtener la ruta de la carpeta donde está la hoja activa
    rutaBase = ws.Parent.Path & "\"
    
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
        seccion = r.Range.Cells(1, tbl.ListColumns("Sección").Index).value
        descripcion = r.Range.Cells(1, tbl.ListColumns("Descripción").Index).value
        imagen = r.Range.Cells(1, tbl.ListColumns("Imágenes").Index).value
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
        
        ' Agregar un párrafo con la descripción
        If Trim(descripcion) <> "" Then
            With docWord.content.Paragraphs.Add
                .Range.text = descripcion
                .Range.Style = docWord.Styles("Normal")        ' Aplicar un estilo predeterminado para el párrafo de descripción
                .Format.SpaceBefore = 12        ' Espacio antes del párrafo para separación
            End With
        End If
        
        docWord.content.InsertParagraphAfter
        docWord.content.Paragraphs.Last.Range.Select
        
        ' Agregar la imagen si existe
        If imagen <> "" Then
            ' Verificar si la imagen existe
            If Dir(imagenRutaCompleta) <> "" Then
                ' Agregar un párrafo vacío para la imagen
                docWord.content.InsertParagraphAfter
                Set rng = docWord.content.Paragraphs.Last.Range
                
                ' Insertar la imagen en el párrafo vacío
                Set shape = docWord.InlineShapes.AddPicture(fileName:=imagenRutaCompleta, LinkToFile:=False, SaveWithDocument:=True, Range:=rng)
                
                ' Centrar la imagen
                shape.Range.ParagraphFormat.Alignment = 1        ' 1 = wdAlignParagraphCenter
                
                ' Agregar un párrafo vacío después de la imagen para el caption
                docWord.content.InsertParagraphAfter
                Set rng = docWord.content.Paragraphs.Last.Range
                
                ' Insertar el caption debajo de la imagen
                Set captionRange = rng.Duplicate
                captionRange.Select
                appWord.Selection.MoveLeft Unit:=1, Count:=1, Extend:=0        ' wdCharacter
                appWord.CaptionLabels.Add Name:="Imagen"
                appWord.Selection.InsertCaption Label:="Imagen", TitleAutoText:="InsertarTítulo1", _
                                                Title:="", Position:=1        ' wdCaptionPositionBelow, ExcludeLabel:=0
                appWord.Selection.ParagraphFormat.Alignment = 1        ' wdAlignParagraphCenter
                
                docWord.content.InsertAfter text:=" " & seccion
                
                ' Agregar un párrafo vacío después del caption para separación
                docWord.content.InsertParagraphAfter
                
            Else
                MsgBox "La imagen        '" & imagenRutaCompleta & "' no se encuentra.", vbExclamation
            End If
        End If
        
        ' Agregar el párrafo de resultados si no está vacío
        If Trim(parrafoResultados) <> "" Then
            With docWord.content.Paragraphs.Add
                .Range.text = parrafoResultados
                .Range.Style = docWord.Styles("Normal")        ' Aplicar un estilo predeterminado para el párrafo de resultados
                .Format.SpaceBefore = 12        ' Espacio antes del párrafo para separación
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
            ' Convierte el contenido en un array separado por el carácter de nueva línea
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
            
            ' Convertir la colección de claves en un array
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
            
            ' Convierte el array nuevamente en una cadena concatenada por el carácter de nueva línea
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
    
    ' Verificar que el índice del gráfico es válido
    If graficoIndex < 1 Or graficoIndex > WordDoc.InlineShapes.Count Then
        MsgBox "Índice de gráfico fuera de rango."
        ActualizarGraficoSegunDicionario = False
        Exit Function
    End If
    
    ' Obtener el InlineShape correspondiente al índice
    Set ils = WordDoc.InlineShapes(graficoIndex)
    
    If ils.Type = 12 And ils.HasChart Then
        Set Chart = ils.Chart
        If Not Chart Is Nothing Then
            ' Activar el libro de trabajo asociado al gráfico
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
                    
                    ' Construir el rango dinámico como una cadena
                    sheetIndex = 1
                    dataRangeAddress = CStr(ChartWorkbook.Sheets(sheetIndex).Name & "$A$1:$B$" & CStr(categoryRow - 1))
                    Debug.Print dataRangeAddress
                    
                    ' Actualizar el gráfico con el nuevo rango de datos
                    On Error Resume Next
                    ChartWorkbook.Sheets(sheetIndex).ChartObjects(1).Chart.SetSourceData Source:=Range(dataRangeAddress)
                    If Err.Number <> 0 Then
                        MsgBox "Error al establecer el rango de datos: " & Err.Description
                        Err.Clear
                        ActualizarGraficoSegunDicionario = False
                        Exit Function
                    End If
                    On Error GoTo 0
                    
                    ' Actualizar el gráfico
                    On Error Resume Next
                    Chart.Refresh
                    If Err.Number <> 0 Then
                        MsgBox "Error al actualizar el gráfico: " & Err.Description
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
            MsgBox "El InlineShape seleccionado no contiene un gráfico válido."
            ActualizarGraficoSegunDicionario = False
        End If
    Else
        MsgBox "El InlineShape seleccionado no contiene un gráfico."
        ActualizarGraficoSegunDicionario = False
    End If
    
    Exit Function
    
ErrorHandler:
    MsgBox "Ocurrió un error: " & Err.Description, vbCritical
    ActualizarGraficoSegunDicionario = False
End Function

Sub CYB006_GenerarDocumentosVulnerabilidiadesWord()
    Dim rng As Range
    Dim tbl As ListObject
    Dim WordApp As Object
    Dim WordDoc As Object
    Dim templatePath As String
    Dim outputPath As String
    Dim replaceDic As Object
    Dim cell As Range
    Dim colIndex As Integer
    Dim rowCount As Integer
    Dim i As Integer
    Dim tempFolder As String
    Dim tempFolderPath As String
    Dim saveFolder As String
    Dim selectedRange As Range        ' Variable para almacenar el rango seleccionado por el usuario
    Dim documentsList() As String        ' Lista para almacenar los documentos generados
    
    ' Solicita al usuario seleccionar el rango de celdas que contienen las columnas a considerar
    On Error Resume Next
    Set selectedRange = Application.InputBox("Seleccione el rango de celdas que contienen las columnas a considerar", Type:=8)
    On Error GoTo 0
    
    If selectedRange Is Nothing Then
        MsgBox "No se ha seleccionado un rango válido.", vbExclamation
        Exit Sub
    End If
    
    ' Verifica si el rango seleccionado está dentro de una tabla
    On Error Resume Next
    Set rng = selectedRange.ListObject.Range
    On Error GoTo 0
    
    If rng Is Nothing Then
        MsgBox "El rango seleccionado no está dentro de una tabla.", vbExclamation
        Exit Sub
    End If
    
    ' Solicita al usuario la ruta del documento de Word
    templatePath = Application.GetOpenFilename("Documentos de Word (*.docx), *.docx", , "Seleccione un documento de Word como plantilla")
    If templatePath = "Falso" Then Exit Sub
    
    ' Solicita al usuario la carpeta donde desea guardar los archivos generados
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Seleccione la carpeta donde desea guardar los documentos generados"
        If .Show = -1 Then
            saveFolder = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With
    
    ' Crea una instancia de la aplicación de Word
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
    Dim fs As Object
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
        fs.CopyFile templatePath, tempFolder & "\Documento_" & i & ".docx"
        ' Abre la copia del documento de Word
        Set WordDoc = WordApp.Documents.Open(tempFolder & "\Documento_" & i & ".docx")
        ' Realiza los reemplazos en el documento de Word
        For Each key In replaceDic.Keys
        
            Debug.Print CStr(key)
            If CStr(key) = "«Descripción»" Then
                ' Aplicar la función específica para la clave «Descripcion»
                replaceDic(key) = TransformText(replaceDic(key))
            End If
            If CStr(key) = "Propuesta de remediación" Then
                ' Aplicar la función específica para la clave «Descripcion»
                replaceDic(key) = TransformText(replaceDic(key))
            End If
            If CStr(key) = "Referencias" Then
                ' Aplicar la función específica para la clave «Descripcion»
                replaceDic(key) = TransformText(replaceDic(key))
            End If
            ' Reemplazar en el documento de Word
            WordAppReplaceParagraph WordApp, WordDoc, CStr(key), CStr(replaceDic(key))
        Next key
        FormatRiskLevelCell WordDoc.Tables(1).cell(1, 2)
        ' Guarda y cierra el documento de Word
        ' Antes de guardar el documento de Word
        EliminarUltimasFilasSiEsSalidaPruebaSeguridad WordDoc, replaceDic
        WordDoc.Save
        WordDoc.Close
        
        ' Agregar el documento generado a la lista
        ReDim Preserve documentsList(i - 2)
        documentsList(i - 2) = tempFolder & "\Documento_" & i & ".docx"
    Next i
    
    ' Combina todos los archivos en uno solo
    Dim finalDocumentPath As String
    finalDocumentPath = saveFolder & "\Documento_Consolidado.docx"
    MergeDocuments WordApp, documentsList, finalDocumentPath
    
    ' Mueve la carpeta temporal a la carpeta seleccionada por el usuario
    fs.MoveFolder tempFolder, saveFolder & "\DocumentosGenerados"
    
    ' Cerrar la aplicación de Word
    WordApp.Quit
    Set WordApp = Nothing
    
    ' Muestra un mensaje de éxito
    MsgBox "Se han generado los documentos de Word correctamente.", vbInformation
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
    
    ' Crear un diálogo para seleccionar el archivo CSV
    campoArchivoPath = Application.GetOpenFilename("Archivos CSV (*.csv), *.csv", , "Seleccionar archivo CSV")
    If campoArchivoPath = "Falso" Then
        MsgBox "No se seleccionó ningún archivo CSV. La macro se detendrá."
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
    
    ' Extraer el nombre de la Aplicación del diccionario
    If replaceDic.Exists("«Aplicación»") Then
        appName = replaceDic("«Aplicación»")
    Else
        MsgBox "No se encontró el campo        'Aplicación' en el archivo CSV.", vbExclamation
        Exit Sub
    End If
    
    ' Crear diálogos para seleccionar plantillas y carpeta de salida
    Set dlg = Application.FileDialog(msoFileDialogFilePicker)
    
    dlg.Title = "Seleccionar la plantilla de reporte técnico"
    dlg.Filters.Clear
    dlg.Filters.Add "Archivos de Word", "*.docx"
    If dlg.Show = -1 Then
        plantillaReportePath = dlg.SelectedItems(1)
    Else
        MsgBox "No se seleccionó ningún archivo. La macro se detendrá."
        Exit Sub
    End If
    
    dlg.Title = "Seleccionar la plantilla de reporte ejecutivo"
    If dlg.Show = -1 Then
        plantillaReportePath2 = dlg.SelectedItems(1)
    Else
        MsgBox "No se seleccionó ningún archivo. La macro se detendrá."
        Exit Sub
    End If
    
    dlg.Title = "Seleccionar la plantilla de tabla de vulnerabilidades"
    If dlg.Show = -1 Then
        plantillaVulnerabilidadesPath = dlg.SelectedItems(1)
    Else
        MsgBox "No se seleccionó ningún archivo. La macro se detendrá."
        Exit Sub
    End If
    
    ' Solicitar la carpeta de salida
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Seleccionar Carpeta de Salida"
        If .Show = -1 Then
            carpetaSalida = .SelectedItems(1)
        Else
            MsgBox "No se seleccionó ninguna carpeta. La macro se detendrá."
            Exit Sub
        End If
    End With
    
    ' Crear una subcarpeta con el nombre de la Aplicación
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
    
    ' Crear y abrir documentos de reporte técnico y ejecutivo
    For Each plantilla In Array(plantillaReportePath, plantillaReportePath2)
        archivoTemp = tempFolder & "\" & fileSystem.GetFileName(plantilla)
        fileSystem.CopyFile plantilla, archivoTemp
        
        Set WordDoc = WordApp.Documents.Open(archivoTemp)
        WordApp.Visible = False
        
        ReplaceFields WordDoc, replaceDic
        
        If plantilla = plantillaReportePath Then
            tempDocPath = tempFolder & "\SSIFO14-03 Informe técnico.docx"
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
        MsgBox "No se ha seleccionado un rango válido.", vbExclamation
        Exit Sub
    End If
    
    Dim resultado As Boolean
    
    ' Llamar a la funcián para exportar la hoja activa a Excel
    resultado = FunExportarHojaActivaAExcelINAI(carpetaSalida, appName)
    
    ' Verifica si el rango seleccionado está dentro de una tabla
    On Error Resume Next
    Set rng = selectedRange.ListObject.Range
    On Error GoTo 0
    
    If rng Is Nothing Then
        MsgBox "El rango seleccionado no está dentro de una tabla.", vbExclamation
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
        MsgBox "No se encontrá la columna        'Severidad' en el rango seleccionado.", vbExclamation
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
        MsgBox "No se encontrá la columna        'Tipo de vulnerabilidad' en el rango seleccionado.", vbExclamation
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
        
        ReplaceFields WordDoc, replaceDic
        FormatRiskLevelCell WordDoc.Tables(1).cell(1, 2)
        EliminarUltimasFilasSiEsSalidaPruebaSeguridad WordDoc, replaceDic
        WordDoc.Save
        WordDoc.Close
        
        numDocuments = numDocuments + 1
        ReDim Preserve documentsList(numDocuments - 1)
        documentsList(numDocuments - 1) = tempFolderGenerados & "\" & tempFileName
    Next i
    
    ' Combina todos los archivos en uno solo
    finalDocumentPath = tempFolder & "\Tablas_vulnerabilidades.docx"
    MergeDocuments WordApp, documentsList, finalDocumentPath
    
    ' Actualizar el documento de reporte técnico
    Set WordDoc = WordApp.Documents.Open(tempDocPath)
    secVulnerabilidades = "{{Sección de tablas de vulnerabilidades}}"
    
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
    
    ' Actualizar el gráfico InlineShape námero 1 en reporte técnico
    FunActualizarGraficoSegunDicionario WordDoc, severityCounts, 1
    
    ' Actualizar todos los gráficos en el documento
    ActualizarGraficos WordDoc
    ' Update the Table of Contents
    On Error Resume Next
    WordDoc.TablesOfContents(1).Update
    On Error GoTo 0
    
    ' Guardar el documento de reporte técnico final en la subcarpeta
    WordDoc.SaveAs2 carpetaSalida & "\SSIFO14-03 Informe técnico.docx"
    
    ' Guardar como PDF
    nombrePDF = carpetaSalida & "\SSIFO14-03 Informe técnico.pdf"
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
    
    ' Actualizar el gráfico InlineShape námero 1 en reporte ejecutivo
    FunActualizarGraficoSegunDicionario WordDoc, severityCounts, 1
    
    ' Actualizar el gráfico InlineShape námero 2 en reporte ejecutivo
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
    
    ' Cerrar la Aplicación de Word
    WordApp.Quit
    Set WordDoc = Nothing
    Set WordApp = Nothing
    Set fileSystem = Nothing
    
    ' Mostrar mensaje de áxito
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
        MsgBox "No se ha seleccionado un rango válido.", vbExclamation
        Exit Sub
    End If
    
    ' Verificar si el rango seleccionado pertenece a una tabla (ListObject)
    If Not selectedRange.ListObject Is Nothing Then
        ' Si es parte de una tabla, obtenemos el rango de la tabla
        Set tableRange = selectedRange.ListObject.Range
    Else
        ' Si no es parte de una tabla, mostrar un mensaje y salir
        MsgBox "El rango seleccionado no está dentro de una tabla.", vbExclamation
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
    
    ' Extraer el nombre de la Aplicación
    If replaceDic.Exists("«Nombre de carpeta»") Then
        folderName = replaceDic("«Nombre de carpeta»")
    Else
        MsgBox "No se encontró el campo        'Nombre de carpeta'.", vbExclamation
        Exit Sub
    End If
    
    ' Crear una subcarpeta con el nombre de la Aplicación
    carpetaSalida = carpetaSalida & "\" & folderName
    On Error Resume Next
    MkDir carpetaSalida
    On Error GoTo 0
    
    If replaceDic.Exists("«Tipo de reporte»") Then
        Select Case replaceDic("«Tipo de reporte»")
            Case "Técnico"
                
                ' Obtener la ruta de la plantilla directamente de la celda de la tabla
                If replaceDic.Exists("«Ruta de la plantilla»") Then
                    plantillaReportePath = replaceDic("«Ruta de la plantilla»")
                Else
                    MsgBox "No se encontró el campo        'Ruta de la plantilla'.", vbExclamation
                    Exit Sub
                End If
                
                ' Verificar que la ruta de la plantilla exista
                If Len(Dir(plantillaReportePath)) = 0 Then
                    MsgBox "La ruta de la plantilla no es válida o el archivo no existe: " & plantillaReportePath, vbExclamation
                    Exit Sub
                End If
                
                Set dlg = Application.FileDialog(msoFileDialogFilePicker)
                ' Crear diálogos para seleccionar la carpeta de salida
                With Application.FileDialog(msoFileDialogFolderPicker)
                    .Title = "Seleccionar Carpeta de Salida"
                    If .Show = -1 Then
                        carpetaSalida = .SelectedItems(1)
                    Else
                        MsgBox "No se seleccionó ninguna carpeta. La macro se detendrá."
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
                    MsgBox "No se seleccionó ningún archivo. La macro se detendrá."
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
                ReplaceFields WordDoc, replaceDic
                finalDocumentPath = carpetaSalida & "\" & "Informe Técnico.docx"
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
                    MsgBox "No se ha seleccionado un rango válido.", vbExclamation
                    Exit Sub
                End If
                
                ' Verificar si el rango seleccionado pertenece a una tabla (ListObject)
                If Not selectedRange.ListObject Is Nothing Then
                    ' Si es parte de una tabla, obtenemos el rango de la tabla
                    Set tableRange = selectedRange.ListObject.Range
                Else
                    ' Si no es parte de una tabla, mostrar un mensaje y salir
                    MsgBox "El rango seleccionado no está dentro de una tabla.", vbExclamation
                    Exit Sub
                End If
                
                Dim resultado As Boolean
                
                ' Llamar a la funcián para exportar la hoja activa a Excel
                resultado = FunExportarHojaActivaAExcelINAI(carpetaSalida, folderName, tableRange.Worksheet, replaceDic("«Nombre del reporte»"))
                
            Case "Tablas de vulnerabilidades"
                
                GenerarDocumentosVulnerabilidiadesWord (replaceDic("«Nombre del reporte»"))
                
            Case Else
                MsgBox "El tipo de reporte no es reconocido.", vbExclamation
                Exit Sub
        End Select
    Else
        MsgBox "No se encontró el campo        'Tipo de reporte'.", vbExclamation
        Exit Sub
    End If
    
    KillAllWordInstances
    
    MsgBox "Proceso completado correctamente."
End Sub

Sub CYB023_KillAllWordInstances()
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

Sub CYB031_ReplaceFields(WordDoc As Object, replaceDic As Object)
    Dim key         As Variant
    Dim WordApp     As Object
    Dim docContent  As Object
    Dim findInRange As Boolean
    
    ' Obtener la aplicación de Word
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
    ' Actualizar todos los gráficos en el documento de Word
    On Error Resume Next
    
    ' Recorrer todos los InlineShapes en el documento
    Dim i           As Integer
    Dim Chart       As Object
    Dim ChartData   As Object
    Dim ChartWorkbook As Object
    
    For i = 1 To WordDoc.InlineShapes.Count
        With WordDoc.InlineShapes(i)
            ' Verificar si el InlineShape es un gráfico (wdInlineShapeChart = 12)
            If .Type = 12 And .HasChart Then
                Set Chart = .Chart
                If Not Chart Is Nothing Then
                    ' Activar los datos del gráfico
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
                        ' Refrescar el gráfico
                        Chart.Refresh
                    End If
                End If
            End If
        End With
    Next i
End Sub

Function CYB022_GenerarDocumentosVulnerabilidiadesWord(fileName As String)
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
        MsgBox "No se ha seleccionado un rango válido.", vbExclamation
        Exit Function
    End If
    
    ' Verifica si el rango seleccionado está dentro de una tabla
    On Error Resume Next
    Set rng = selectedRange.ListObject.Range
    On Error GoTo 0
    
    If rng Is Nothing Then
        MsgBox "El rango seleccionado no está dentro de una tabla.", vbExclamation
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
    
    ' Crea una instancia de la aplicación de Word
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
            If CStr(key) = "«Descripción»" Then
                ' Aplicar la función específica para la clave «Descripcion»
                replaceDic(key) = TransformText(replaceDic(key))
            End If
            If CStr(key) = "«Propuesta de remediación»" Then
                replaceDic(key) = TransformText(replaceDic(key))
            End If
            If CStr(key) = "Referencias" Then
                replaceDic(key) = TransformText(replaceDic(key))
            End If
            ' Reemplazar en el documento de Word
            WordAppReplaceParagraph WordApp, WordDoc, CStr(key), CStr(replaceDic(key))
            
           
        Next key
        FormatRiskLevelCell WordDoc.Tables(1).cell(1, 2)
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
    MergeDocuments WordApp, documentsList, finalDocumentPath
    
    ' Mueve la carpeta temporal a la carpeta seleccionada por el usuario
    fs.MoveFolder tempFolder, saveFolder & "\Documentos_generados"
    
    ' Cerrar la aplicación de Word
    WordApp.Quit
    Set WordApp = Nothing
    
    ' Muestra un mensaje de éxito
    MsgBox "Se han generado los documentos de Word correctamente.", vbInformation
End Function

Sub CYB021_FormatRiskLevelCell(cell As Object)
    Dim cellText    As String
    ' Obtener el texto de la celda y eliminar los caracteres especiales
    cellText = Replace(cell.Range.text, vbCrLf, "")
    cellText = Replace(cellText, vbCr, "")
    cellText = Replace(cellText, vbLf, "")
    cellText = Replace(cellText, Chr(7), "")
    
    ' Realizar la comparación utilizando el texto de la celda sin caracteres especiales
    Select Case cellText
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
End Sub

Function TransformText(text As String) As String
    Dim regEx       As Object
    Set regEx = CreateObject("VBScript.RegExp")
    
    ' Configurar la expresión regular para encontrar saltos de línea o saltos de carro sin un punto antes y no seguidos de paréntesis ni de guión
    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .pattern = "([^.()\r\n-])(?![^(]*\)|[-])[^\S\r\n]*[\r\n]+"        ' Expresión regular para encontrar saltos de línea o saltos de carro sin un punto antes y no seguidos de paréntesis ni de guión
    End With
    
    ' Realizar la transformación: quitar caracteres especiales y aplicar la expresión regular
    TransformText = regEx.Replace(Replace(text, Chr(7), ""), "$1 ")
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
    
    ' Verificar que el ándice del gráfico es válido
    If graficoIndex < 1 Or graficoIndex > WordDoc.InlineShapes.Count Then
        MsgBox "Índice de gráfico fuera de rango."
        FunActualizarGraficoSegunDicionario = False
        Exit Function
    End If
    
    ' Obtener el InlineShape correspondiente al ándice
    Set ils = WordDoc.InlineShapes(graficoIndex)
    
    If ils.Type = 12 And ils.HasChart Then
        Set Chart = ils.Chart
        If Not Chart Is Nothing Then
            ' Activar el libro de trabajo asociado al gráfico
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
                    
                    ' Construir el rango dinámico como una cadena
                    dataRangeAddress =        '" & SourceSheet.Name & "'!$A$1:$B$" & (categoryRow - 1)
                    Debug.Print dataRangeAddress
                    
                    ' Verifica si la tabla existe usando el ándice
                    On Error Resume Next
                    Set DataTable = SourceSheet.ListObjects(tableIndex)        ' Obtiene el objeto de la tabla por ándice
                    On Error GoTo 0
                    
                    ' Verifica que el objeto de la tabla no sea Nothing
                    If Not DataTable Is Nothing Then
                        ' Redimensiona la tabla al nuevo rango usando el objeto Worksheet
                        DataTable.Resize SourceSheet.Range("A1:B" & (categoryRow - 1))
                    Else
                        MsgBox "La tabla en el índice " & tableIndex & " no se encontró en la hoja."
                    End If
                    
                    WordDoc.InlineShapes(graficoIndex).Chart.SetSourceData Source:=dataRangeAddress
                    
                    ' Actualizar el gráfico
                    Chart.Refresh
                    
                    ' Cerrar el libro de trabajo sin guardar cambios
                    ChartWorkbook.Close SaveChanges:=False
                    
                    FunActualizarGraficoSegunDicionario = True
                End If
            End If
        Else
            MsgBox "El InlineShape seleccionado no contiene un gráfico válido."
            FunActualizarGraficoSegunDicionario = False
        End If
    Else
        MsgBox "El InlineShape seleccionado no contiene un gráfico."
        FunActualizarGraficoSegunDicionario = False
    End If
    
    Exit Function
    
ErrorHandler:
    MsgBox "Ocurrió un error: " & Err.Description, vbCritical
    FunActualizarGraficoSegunDicionario = False
End Function




Sub CYB025_MergeDocuments(WordApp As Object, documentsList As Variant, finalDocumentPath As String)
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
        oRng.InsertFile sFile        ' Insertar el contenido del archivo actual
        
        ' Insertar un salto de página después de cada documento insertado (excepto el último)
        If i < UBound(documentsList) Then
            Set oRng = baseDoc.Range
            oRng.Collapse 0        ' Colapsar el rango al final del documento base
            'oRng.InsertBreak Type:=6 ' Insertar un salto de página
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


Sub CYB028_WordAppReplaceParagraph(WordApp As Object, WordDoc As Object, wordToFind As String, replaceWord As String)
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

    ' Verificar si la celda no está vacía
    If IsEmpty(celda) Then
        MsgBox "Seleccione una celda con un rango de IPs.", vbExclamation, "Error"
        Exit Sub
    End If

    ipRango = Trim(celda.value)

    ' Dividir el rango de IPs usando el guion como separador
    partes = Split(ipRango, "-")
    
    ' Verificar que haya dos partes en el rango
    If UBound(partes) <> 1 Then
        MsgBox "Formato inválido. Use: 10.0.1.60-10.0.1.78", vbExclamation, "Error"
        Exit Sub
    End If

    ipInicio = partes(0)
    ipFin = partes(1)

    ' Extraer el último número de las IPs
    numInicio = CInt(Split(ipInicio, ".")(3))
    numFin = CInt(Split(ipFin, ".")(3))

    ' Validar que el inicio es menor o igual que el fin
    If numInicio > numFin Then
        MsgBox "El rango de IPs es inválido.", vbExclamation, "Error"
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

1
