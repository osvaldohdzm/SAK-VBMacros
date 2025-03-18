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
                        If Not uniqueUrls.exists(contentArray(i)) Then
                            uniqueUrls.Add contentArray(i), Nothing
                        End If
                    End If
                End If
            Next i
            
            ' Convertir la colección de claves en un array
            n = uniqueUrls.Count - 1
            ReDim uniqueArray(n)
            i = 0
            For Each key In uniqueUrls.keys
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
            If Not celda.HasFormula Then ' Ignora celdas con fórmulas

                ' Reemplazar tabulaciones con espacios
                celda.value = Replace(celda.value, Chr(9), " ")
                ' Eliminar espacios innecesarios
                celda.value = Application.Trim(celda.value)

                ' Reemplazar <li> por saltos de línea
                celda.value = RegExpReplace(celda.value, liPattern, vbLf)

                ' Eliminar etiquetas HTML dejando solo texto
                celda.value = RegExpReplace(celda.value, cleanHtmlPattern, vbNullString)

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
                Texto = ReplaceHtmlEntities(cleanOutput)

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
    Dim regex       As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    With regex
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .pattern = replacePattern
    End With
    
    RegExpReplace = regex.Replace(text, replaceWith)
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
    
    ' Verificar si la celda seleccionada está dentro de una tabla
    On Error Resume Next
    Set tabla = celdaActual.ListObject
    On Error GoTo 0
    
    If tabla Is Nothing Then
        MsgBox "Debes seleccionar una celda dentro de una tabla para ejecutar la ordenación.", vbExclamation, "Error"
        Exit Sub
    End If
    
    ' Confirmar con el usuario antes de proceder
    respuesta = MsgBox("Se ordenará la tabla por la columna 'Severidad' según el color de relleno." & vbCrLf & _
                       "Orden: Morado ? Rojo ? Amarillo ? Verde." & vbCrLf & vbCrLf & _
                       "¿Deseas continuar?", vbYesNo + vbQuestion, "Confirmación")
    
    If respuesta <> vbYes Then Exit Sub
    
    ' Aplicar ordenación por color en el orden definido
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
    
    MsgBox "Ordenación completada: Morado ? Rojo ? Amarillo ? Verde.", vbInformation, "Proceso finalizado"
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
    
    ' Verificar si las claves están presentes en el diccionario
    If replaceDic.exists(salidaPruebaSeguridadKey) And replaceDic.exists(metodoDeteccionKey) Then
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
                        If Not uniqueUrls.exists(contentArray(i)) Then
                            uniqueUrls.Add contentArray(i), Nothing
                        End If
                    End If
                End If
            Next i
            
            ' Convertir la colección de claves en un array
            n = uniqueUrls.Count - 1
            ReDim uniqueArray(n)
            i = 0
            For Each key In uniqueUrls.keys
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
                    For Each category In conteos.keys
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
        For Each key In replaceDic.keys
        
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
        FormatDashParagraphsCell WordDoc.Tables(1).cell(4, 2)
        FormatDashParagraphsCell WordDoc.Tables(1).cell(5, 2)
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
    If replaceDic.exists("«Aplicación»") Then
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
            If severityCounts.exists(severity) Then
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
            If vulntypesCounts.exists(vulntypes) Then
                vulntypesCounts(vulntypes) = vulntypesCounts(vulntypes) + 1
            Else
                vulntypesCounts.Add vulntypes, 1
            End If
        End If
    Next i
    
    ' Inicializar conteos
    countBAJA = IIf(severityCounts.exists("BAJA"), severityCounts("BAJA"), 0)
    countMEDIA = IIf(severityCounts.exists("MEDIA"), severityCounts("MEDIA"), 0)
    countALTA = IIf(severityCounts.exists("ALTA"), severityCounts("ALTA"), 0)
    countCRITICAS = IIf(severityCounts.exists("CRÍTICOS"), severityCounts("CRÍTICOS"), 0)
    
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
        If replaceDic.exists(key) Then
            MsgBox "Se ha encontrado un encabezado duplicado: " & headerRow.Cells(1, i).value & _
                   vbCrLf & "Por favor, corrige los encabezados duplicados y vuelve a ejecutar la macro.", vbExclamation
            Exit Sub
        End If
        
        ' Asignar el valor de la celda de la fila seleccionada al diccionario
        value = selectedRange.Cells(1, i).value
        replaceDic.Add key, value
    Next i
    
    ' Extraer el nombre de la Aplicación
    If replaceDic.exists("«Nombre de carpeta»") Then
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
    
    If replaceDic.exists("«Tipo de reporte»") Then
        Select Case replaceDic("«Tipo de reporte»")
            Case "Técnico"
                
                ' Obtener la ruta de la plantilla directamente de la celda de la tabla
                If replaceDic.exists("«Ruta de la plantilla»") Then
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

Sub ReplaceFields(WordDoc As Object, replaceDic As Object)
    Dim key         As Variant
    Dim WordApp     As Object
    Dim docContent  As Object
    Dim findInRange As Boolean
    
    ' Obtener la aplicación de Word
    Set WordApp = WordDoc.Application
    
    ' Obtener el contenido del documento
    Set docContent = WordDoc.content
    
    ' Bucle para buscar y reemplazar todas las ocurrencias en el diccionario
    For Each key In replaceDic.keys
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
        For Each key In replaceDic.keys
            
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



Sub FormatDashParagraphsCell(cell As Object)
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

    ' Recorre cada párrafo dentro de la celda
    For Each p In cell.Range.Paragraphs
        strTexto = p.Range.text
        
        ' Si el párrafo comienza con "- "
        If Left(Trim(strTexto), 2) = "- " Then
            posDosPuntos = InStr(strTexto, ":")
            
            ' Si hay dos puntos en el texto
            If posDosPuntos > 0 Then
                ' Aplicar negrita al texto antes de los dos puntos
                Set rng = p.Range
                rng.Start = p.Range.Start
                rng.End = p.Range.Start + posDosPuntos - 1
                rng.Font.Bold = True
                
                ' El texto después de los dos puntos no tendrá negrita
                Set rng = p.Range
                rng.Start = p.Range.Start + posDosPuntos
                rng.End = p.Range.End
                rng.Font.Bold = False
            End If
        End If
    Next p
End Sub

Sub FormatRiskLevelCell(cell As Object)
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
        ' Si no es un número, usar la clasificación por texto
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


Function TransformText(text As String) As String
    Dim regex       As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    ' Configurar la expresión regular para encontrar saltos de línea o saltos de carro sin un punto antes y no seguidos de paréntesis ni de guión
    With regex
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .pattern = "([^.()\r\n-])(?![^(]*\)|[-])[^\S\r\n]*[\r\n]+"        ' Expresión regular para encontrar saltos de línea o saltos de carro sin un punto antes y no seguidos de paréntesis ni de guión
    End With
    
    ' Realizar la transformación: quitar caracteres especiales y aplicar la expresión regular
    TransformText = regex.Replace(Replace(text, Chr(7), ""), "$1 ")
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
                    For Each category In conteos.keys
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




Sub MergeDocuments(WordApp As Object, documentsList As Variant, finalDocumentPath As String)
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


Sub WordAppReplaceParagraph(WordApp As Object, WordDoc As Object, wordToFind As String, replaceWord As String)
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




Sub CYB034_CargarDatosDesdeCSVNessus()
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
    
    ' Verificar si la celda activa está dentro de una tabla
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
    
    ' Si no está dentro de una tabla, salir
    If Not celdaEnTabla Then
        MsgBox "La celda seleccionada no está dentro de una tabla.", vbExclamation
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
                    Case "Description": columnaCorrespondiente = "Descripción ampliada"
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
    
    MsgBox "Datos cargados con éxito en la tabla.", vbInformation
End Sub


Sub CYB035_CargarDatosDesdeCSVNexPose()
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
    
    ' Verificar si la celda activa está dentro de una tabla
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
    
    ' Si no está dentro de una tabla, salir
    If Not celdaEnTabla Then
        MsgBox "La celda seleccionada no está dentro de una tabla.", vbExclamation
        Exit Sub
    End If
    
    ' Preguntar al usuario si está seguro de cargar los datos
    mensaje = "¿Está seguro que desea cargar datos de los archivos CSV en la tabla '" & tbl.Name & "'?"
    respuesta = MsgBox(mensaje, vbYesNo + vbQuestion, "Confirmación")
    If respuesta = vbNo Then Exit Sub
    
    ' Seleccionar múltiples archivos CSV
    archivos = Application.GetOpenFilename("Archivos CSV (*.csv), *.csv", MultiSelect:=True, Title:="Seleccionar archivos CSV")
    If Not IsArray(archivos) Then Exit Sub ' Si el usuario cancela la selección
    
    ' Iterar sobre los archivos seleccionados
    For Each archivo In archivos
        ' Verificar que el archivo seleccionado es CSV
        If LCase(Trim(Right(archivo, 4))) <> ".csv" Then
            MsgBox "El archivo " & archivo & " no es un archivo CSV.", vbExclamation
            Exit Sub
        End If
        
        ' Abrir el archivo CSV
        Set wbCSV = Workbooks.Open(fileName:=archivo, Local:=True)
        Set wsCSV = wbCSV.Sheets(1) ' Asumimos que los datos están en la primera hoja
        
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
                        tbl.ListColumns("Identificador de detección usado").DataBodyRange.Cells(tbl.ListRows.Count, 1).value = csvData(i, colCSVIndex)
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
    
    MsgBox "Datos cargados con éxito en la tabla.", vbInformation
End Sub


Sub CYB036_CargarDatosDesdeXMLOpenVAS()
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
    
    ' Verificar si la celda activa está dentro de alguna tabla y asignarla a tbl
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
        MsgBox "La celda seleccionada no está dentro de una tabla.", vbExclamation
        Exit Sub
    End If
    
    ' Confirmar con el usuario
    mensaje = "¿Está seguro que desea cargar datos del archivo XML en la tabla '" & tbl.Name & "'?"
    respuesta = MsgBox(mensaje, vbYesNo + vbQuestion, "Confirmación")
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
        If Not dict.exists(field) Then
            MsgBox "La columna '" & field & "' no se encontró en la tabla.", vbExclamation
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

    MsgBox "Datos cargados con éxito.", vbInformation
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
        MsgBox "No se encontró la tabla 'Tbl_falses_positives' en ninguna hoja.", vbExclamation
        Exit Sub
    End If

    ' Obtener el índice de la columna "Vulnerability Name"
    On Error Resume Next
    columnaIndex = tbl.ListColumns("Vulnerability Name").Index
    On Error GoTo 0
    If columnaIndex = 0 Then
        MsgBox "La columna 'Vulnerability Name' no se encontró en la tabla.", vbExclamation
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
        If valoresTabla.exists(celda.value) Then
            celda.Interior.Color = RGB(0, 255, 0) ' Verde chillón
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
Attribute CYB041_IrACatalogoVulnerabilidad.VB_ProcData.VB_Invoke_Func = "G\n14"
   
   Dim wsOrigen As Worksheet, wsCatalogo As Worksheet
    Dim tblOrigen As ListObject, tblCatalogo As ListObject
    Dim rngCeldaActual As Range
    Dim idVulnerabilidad As Variant, tipoOrigen As String
    Dim colBusqueda As String, rngBusqueda As Range, celdaEncontrada As Range
    Dim dictColumnas As Object
    
    ' Definir las hojas
    Set wsOrigen = ActiveSheet ' La hoja actual donde se ejecuta la macro
    Set wsCatalogo = ThisWorkbook.Sheets("Catalogo vulnerabilidades") ' Hoja donde está la tabla de catálogo
    
    ' Identificar la tabla en la hoja actual (se asume que solo hay una tabla)
    If wsOrigen.ListObjects.Count = 0 Then
        MsgBox "No se encontró una tabla en la hoja actual.", vbExclamation, "Error"
        Exit Sub
    End If
    Set tblOrigen = wsOrigen.ListObjects(1) ' Toma la primera tabla de la hoja actual
    
    ' Crear diccionario con los nombres de columnas según el tipo de origen
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

    ' Obtener la celda actual
    Set rngCeldaActual = ActiveCell
    
    ' Verificar si la celda actual está dentro de la tabla
    If Intersect(rngCeldaActual, tblOrigen.DataBodyRange) Is Nothing Then
        MsgBox "Por favor, selecciona una celda dentro de la tabla de vulnerabilidades.", vbExclamation, "Error"
        Exit Sub
    End If
    
    ' Obtener las columnas "Tipo de origen" e "Identificador original de la vulnerabilidad"
    On Error Resume Next
    tipoOrigen = tblOrigen.ListColumns("Tipo de origen").DataBodyRange.Cells(rngCeldaActual.Row - tblOrigen.DataBodyRange.Row + 1, 1).value
    idVulnerabilidad = tblOrigen.ListColumns("Identificador original de la vulnerabilidad").DataBodyRange.Cells(rngCeldaActual.Row - tblOrigen.DataBodyRange.Row + 1, 1).value
    On Error GoTo 0
    
    ' Validar si se obtuvo un tipo de origen y un ID válido
    If tipoOrigen = "" Or IsEmpty(idVulnerabilidad) Then
        MsgBox "No se encontró un Tipo de Origen o Identificador válido en la fila actual.", vbExclamation, "Error"
        Exit Sub
    End If
    
    ' Verificar si el tipo de origen está en el diccionario
    If Not dictColumnas.exists(tipoOrigen) Then
        MsgBox "El tipo de origen '" & tipoOrigen & "' no tiene una columna asignada en la tabla de catálogo.", vbExclamation, "Error"
        Exit Sub
    End If

    ' Obtener el nombre de la columna de búsqueda en el catálogo
    colBusqueda = dictColumnas(tipoOrigen)

    ' Obtener la tabla de catálogo
    On Error Resume Next
    Set tblCatalogo = wsCatalogo.ListObjects("Tbl_Catalogo_vulnerabilidades")
    On Error GoTo 0

    If tblCatalogo Is Nothing Then
        MsgBox "La tabla 'Tbl_Catalogo_vulnerabilidades' no existe en la hoja 'Catalogo vulnerabilidades'.", vbExclamation, "Error"
        Exit Sub
    End If

    ' Buscar la columna correspondiente en el catálogo
    Set rngBusqueda = tblCatalogo.ListColumns(colBusqueda).DataBodyRange

    ' Buscar el identificador en la columna correspondiente
    Set celdaEncontrada = rngBusqueda.Find(What:=idVulnerabilidad, LookAt:=xlWhole)

    ' Si se encontró, seleccionar la fila correspondiente en el catálogo
    If Not celdaEncontrada Is Nothing Then
        wsCatalogo.Activate
        celdaEncontrada.EntireRow.Select
        MsgBox "Registro encontrado. Se ha seleccionado la fila correspondiente en el catálogo.", vbInformation, "Éxito"
    Else
        MsgBox "No se encontró el identificador en la tabla de catálogo.", vbExclamation, "Registro no encontrado"
    End If
End Sub


Sub CYB042_Standardize()
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
            If Not dict.exists(key) Then
                dict.Add key, CreateObject("Scripting.Dictionary")
            End If
            
            ' Guardar valores no vacíos en cada columna relevante
            For Each colName In colIndex.keys
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
        If key <> "" And dict.exists(key) Then
            For Each colName In colIndex.keys
                colNum = colIndex(colName)
                If ws.Cells(i, colNum).value = "" And dict(key).exists(colName) Then
                    ws.Cells(i, colNum).value = dict(key)(colName)
                End If
            Next
        End If
    Next i
    
    MsgBox "Estandarización completada.", vbInformation
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
    
    ' Recorrer cada celda en la selección
    For Each cell In Selection
        ' Verificar si la celda no está vacía
        If Not IsEmpty(cell.value) Then
            Vulnerabilidad = cell.value
            
            ' Construcción del prompt
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
    
    ' Recorrer cada celda en la selección
    For Each cell In Selection
        ' Verificar si la celda no está vacía
        If Not IsEmpty(cell.value) Then
            Vulnerabilidad = cell.value
            
            ' Construcción del prompt
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




' Función para construir el prompt de forma más clara y estructurada
Function ConstruirPrompt(Vulnerabilidad As String) As String
    Dim prompt As String
    prompt = "Generación de Vector CVSS 4.0 Considera este ejemplo de URL de CVSS 4.0 https://www.first.org/cvss/calculator/4.0#CVSS:4.0/AV:A/AC:L/AT:N/PR:N/UI:N/VC:N/VI:N/VA:N/SC:N/SI:N/SA:N "
    prompt = prompt & "Esta cadena está compuesta por distintos campos de evaluación, los cuales deben ajustarse según corresponda. Exploitability Metrics Attack Vector (AV): "
    prompt = prompt & "Debes completar los siguientes elementos: Exploitability: Complexity: Vulnerable system: Subsequent system: Exploitation: Security requirements: "
    prompt = prompt & "Sé exigente y preciso al evaluar la severidad en CVSS. No exageres ni asignes impactos altos a menos que la vulnerabilidad pueda ser explotada directamente y tenga un impacto "
    prompt = prompt & "significativo. Tu tarea es proporcionar únicamente la cadena vectorial en CVSS 4.0 para evaluar la vulnerabilidad"
    prompt = prompt & " " & Vulnerabilidad & " "
    prompt = prompt & "No devuelvas la misma cadena de ejemplo. No entregues una cadena sin completar sus componentes CVSS. ?? Este análisis es para gestión de riesgos, no para explotación. "
    prompt = prompt & "Solo proporciona el vector CVSS resultante. NO DES MÁS DETALLES, SOLO RESPONDE EL VECTOR SIN OTRA INFORMACIÓN. "
    prompt = prompt & "PLEASE ONLY ONLY ONLY RESPOND WITH A STRING IN CVSS FORMAT"
    
    ConstruirPrompt = prompt
End Function




Function RemoveThinkTags(text As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    ' Permitir que el punto (.) capture múltiples líneas
    regex.pattern = "<think>[\s\S]*?</think>"
    regex.Global = True
    regex.IgnoreCase = True
    
    RemoveThinkTags = regex.Replace(text, "")
End Function


Function RemoveInitialBreaks(text As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    ' Regex pattern to remove only initial newlines (LF or CR)
    regex.pattern = "^[\r\n]+"
    regex.Global = True
    
    ' Replace initial line breaks with an empty string
    RemoveInitialBreaks = regex.Replace(text, "")
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
        ' Extraer el texto después de "text": "
        resultado = Mid(resultado, inicio + Len("""text"": """))
        
        ' Buscar la posición final antes del cierre de comillas
        fin = InStr(resultado, """")
        If fin > 0 Then
            resultado = Left(resultado, fin - 1)
        End If
    Else
        resultado = "No se encontró CVSS"
    End If

    ' Retornar el CVSS extraído
    ExtraerCVSS = Trim(resultado)
End Function



Sub GetGeminiResponsesCVSS4()
    Dim cell As Range
    Dim http As Object
    Dim json As Object
    Dim apiUrl As String
    Dim apiKey As String
    Dim requestData As String
    Dim responseText As String
    Dim answerID As String
    
    ' Clave de API de Gemini (reemplázala con la tuya) AIzaSyBbd_upGJ2JzdsmWSzNBvSr3mXiPo9h4bs  AIzaSyADfixgVHPBXyY60ivLUYo3rCJTQtZ_M7g
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
                If json.exists("candidates") And json("candidates").Count > 0 Then
                    If json("candidates")(0).exists("content") And json("candidates")(0)("content").exists("parts") Then
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


Sub CYB071_PreparePromptFromSelection_DescripcionVuln_OverVPN()
    Dim celda As Range
    Dim listaVulnerabilidades As String
    Dim prompt As String
    
    ' Inicializar la variable que almacenará las vulnerabilidades
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
        prompt = "Redacta un párrafo técnico breve y conciso que describa la vulnerabilidad detectada, comenzando con la frase: -El sistema…-. Explica en qué consiste la debilidad de seguridad de manera técnica. No incluyas escenarios de explotación, ya que eso corresponde a otro campo. No describas cómo se explota, solo en qué consiste el problema. No menciones el nombre exacto de la vulnerabilidad; utiliza expresiones similares. Vulnerabilidades a describi en formato tabla: "
        prompt = prompt & " " & vbCrLf & Chr(10) & listaVulnerabilidades & vbCrLf & Chr(10)
        prompt = prompt & "Contexto de análisis: Análisis de vulnerabilidades de infraestructura a partir conexion a red VPN. Vulnerabilidad detectada mediante escaneo e interacciones desde red privada VPN."

        
        ' Copiar el prompt generado al portapapeles
        CopyToClipboard prompt
        
        ' Mostrar un mensaje informativo
        MsgBox "El prompt ha sido copiado al portapapeles.", vbInformation, "Prompt Generado"
    Else
        MsgBox "No se encontraron valores en la selección.", vbExclamation, "Sin Datos"
    End If
End Sub



Sub CYB072_PreparePromptFromSelection_AmenazaVuln_OverVPN()
    Dim celda As Range
    Dim listaVulnerabilidades As String
    Dim prompt As String
    
    ' Inicializar la variable que almacenará las vulnerabilidades
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
        prompt = "Hola, por favor considera el siguiente ejemplo: Un atacante podría (obtener, realizar, ejecutar, visualizar, identificar, listar)… Algunos posibles vectores de ataque adicionales asociados con esta amenaza incluyen (pueden ser algunos de los siguientes, pero no necesariamente todos): •  Malware: Un malware diseñado para automatizar intentos de fuerza bruta podría explotar la vulnerabilidad para… •    Usuario malintencionado: Un usuario dentro de la red local con conocimiento de la vulnerabilidad podría aprovecharla para… •    Personal interno: Un empleado con acceso y conocimientos técnicos podría, intencionalmente o por error,… •  Delincuente cibernético: Un atacante externo en busca de vulnerabilidades podría intentar explotar esta debilidad para… Instrucciones adicionales: 1.   Pregunta si el sistema es interno o externo para determinar los vectores de ataque más relevantes, ya que no todos aplican en todos los casos. "
        prompt = prompt & "2.   Redacta una descripción de la amenaza que incluya posibles escenarios de ataque, comenzando con la frase: -Un atacante podría…-. 3.    No es necesario proporcionar un ejemplo para cada vector de ataque. Selecciona el más realista o probable. 4.   En las viñetas de vectores de ataque, menciona la probabilidad de ocurrencia, incluso si es baja. 5.    El contexto es un análisis de vulnerabilidades de infraestructura desde una red privada, específicamente: -Vulnerabilidad detectada mediante escaneo e interacciones desde red privada-. Formato de respuesta: •    Responde en una tabla de dos columnas. •    Para cada vulnerabilidad, redacta un párrafo descriptivo en la primera columna (mínimo 75 palabras). •  En la segunda columna, lista los vectores de ataque con viñetas (usando guiones  - ). • No uses HTML, solo texto plano. Ejemplo de estructura: Descripción de la amenaza    Vectores de ataque Un atacante podría explotar esta vulnerabilidad para acceder a información"
        prompt = prompt & "confidencial Esta amenaza es particularmente crítica en sistemas internos donde los controles de seguridad son menos estrictos."
        prompt = prompt & "Un escenario probable incluye...    - Malware: Un malware podría ser utilizado para... (probabilidad media). - Usuario malintencionado: Un empleado con acceso podría... (probabilidad baja). - Delincuente cibernético: Un atacante externo podría... (probabilidad alta). ES MUY IMPORANTE QU EPAR ALOS VECTORI DE ATQUE DE LA MENAZA USSES GUIONE S MEDIOS COMO VIÑETAS DENTRO DE ALS CELDAS"
        prompt = prompt & "Se detecto mediante VPN red privada, pero explica los escenarios que consideres necesarios" & vbCrLf & Chr(10)
        prompt = prompt & "SOLO DOS columnas, nombre y amenaza"
        prompt = prompt & " " & vbCrLf & Chr(10) & listaVulnerabilidades & vbCrLf & Chr(10)
          
        ' Copiar el prompt generado al portapapeles
        CopyToClipboard prompt
        
        ' Mostrar un mensaje informativo
        MsgBox "El prompt ha sido copiado al portapapeles.", vbInformation, "Prompt Generado"
    Else
        MsgBox "No se encontraron valores en la selección.", vbExclamation, "Sin Datos"
    End If
End Sub


Sub CYB073_PreparePromptFromSelection_PropuestaRemediacionVuln_OverVPN()
    Dim celda As Range
    Dim listaVulnerabilidades As String
    Dim prompt As String
    
    ' Inicializar la variable que almacenará las vulnerabilidades
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
        prompt = "Hola, por favor Redacta como un pentester un párrafo técnico de propuesta de remediación que comience con la frase: -Se recomienda…-. Incluye tantos detalles puntuales par remaicion por jemplo nombre de soluciones que funeriona como poroteccon o controles de sguridad, disoitivos, practicas, menciona de manera más puntual que se pod´ria hacer para que le encargado del sistema, activo, pueda saber como remediari La respueste debe tener SOLO parrafo breve introducción y luego viñetas para lso putnos de propueta de remedaicion Responde para la siguientel lista de vulnerabilidades en FORMATO TABLA DOS COLUMNAS, , SIEMPRE COMIENZAX CON -Se recomienda…- TEXTO AMPLIDO MAS DE 80 palabras, atalmente aplicable a muchos casos, lengauje so framworks"
           prompt = prompt & "Se detecto mediante VPN red privada, pero explica los escenarios que consideres necesarios" & vbCrLf & Chr(10)
          prompt = prompt & "Menciona solo soluciones corporativas"
          prompt = prompt & "Solo dos columnas, nombre y propuesat remediacion"
        prompt = prompt & " " & vbCrLf & Chr(10) & listaVulnerabilidades & vbCrLf & Chr(10)
          
        ' Copiar el prompt generado al portapapeles
        CopyToClipboard prompt
        
        ' Mostrar un mensaje informativo
        MsgBox "El prompt ha sido copiado al portapapeles.", vbInformation, "Prompt Generado"
    Else
        MsgBox "No se encontraron valores en la selección.", vbExclamation, "Sin Datos"
    End If
End Sub


Sub CYB074_PreparePromptFromSelection_DescripcionVuln_OnPrivateNetwork()
    Dim celda As Range
    Dim listaVulnerabilidades As String
    Dim prompt As String
    
    ' Inicializar la variable que almacenará las vulnerabilidades
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
        prompt = "Redacta un párrafo técnico breve y conciso que describa la vulnerabilidad detectada, comenzando con la frase: -El sistema…-. Explica en qué consiste la debilidad de seguridad de manera técnica. No incluyas escenarios de explotación, ya que eso corresponde a otro campo. No describas cómo se explota, solo en qué consiste el problema. No menciones el nombre exacto de la vulnerabilidad; utiliza expresiones similares. Vulnerabilidades a describi en formato tabla: "
        prompt = prompt & " " & vbCrLf & Chr(10) & listaVulnerabilidades & vbCrLf & Chr(10)
        prompt = prompt & "Contexto de análisis: Análisis de vulnerabilidades de infraestructura a partir conexion a red en red privada en sitio. Vulnerabilidad detectada mediante escaneo e interacciones desde red privada en red privada en sitio."

        
        ' Copiar el prompt generado al portapapeles
        CopyToClipboard prompt
        
        ' Mostrar un mensaje informativo
        MsgBox "El prompt ha sido copiado al portapapeles.", vbInformation, "Prompt Generado"
    Else
        MsgBox "No se encontraron valores en la selección.", vbExclamation, "Sin Datos"
    End If
End Sub



Sub CYB075_PreparePromptFromSelection_AmenazaVuln_OnPrivateNetwork()
    Dim celda As Range
    Dim listaVulnerabilidades As String
    Dim prompt As String
    
    ' Inicializar la variable que almacenará las vulnerabilidades
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
        prompt = "Hola, por favor considera el siguiente ejemplo: Un atacante podría (obtener, realizar, ejecutar, visualizar, identificar, listar)… Algunos posibles vectores de ataque adicionales asociados con esta amenaza incluyen (pueden ser algunos de los siguientes, pero no necesariamente todos): •  Malware: Un malware diseñado para automatizar intentos de fuerza bruta podría explotar la vulnerabilidad para… •    Usuario malintencionado: Un usuario dentro de la red local con conocimiento de la vulnerabilidad podría aprovecharla para… •    Personal interno: Un empleado con acceso y conocimientos técnicos podría, intencionalmente o por error,… •  Delincuente cibernético: Un atacante externo en busca de vulnerabilidades podría intentar explotar esta debilidad para… Instrucciones adicionales: 1.   Pregunta si el sistema es interno o externo para determinar los vectores de ataque más relevantes, ya que no todos aplican en todos los casos. "
        prompt = prompt & "2.   Redacta una descripción de la amenaza que incluya posibles escenarios de ataque, comenzando con la frase: -Un atacante podría…-. 3.    No es necesario proporcionar un ejemplo para cada vector de ataque. Selecciona el más realista o probable. 4.   En las viñetas de vectores de ataque, menciona la probabilidad de ocurrencia, incluso si es baja. 5.    El contexto es un análisis de vulnerabilidades de infraestructura desde una red privada, específicamente: -Vulnerabilidad detectada mediante escaneo e interacciones desde red privada-. Formato de respuesta: •    Responde en una tabla de dos columnas. •    Para cada vulnerabilidad, redacta un párrafo descriptivo en la primera columna (mínimo 75 palabras). •  En la segunda columna, lista los vectores de ataque con viñetas (usando guiones  - ). • No uses HTML, solo texto plano. Ejemplo de estructura: Descripción de la amenaza    Vectores de ataque Un atacante podría explotar esta vulnerabilidad para acceder a información"
        prompt = prompt & "confidencial Esta amenaza es particularmente crítica en sistemas internos donde los controles de seguridad son menos estrictos."
        prompt = prompt & "Un escenario probable incluye...    - Malware: Un malware podría ser utilizado para... (probabilidad media). - Usuario malintencionado: Un empleado con acceso podría... (probabilidad baja). - Delincuente cibernético: Un atacante externo podría... (probabilidad alta). ES MUY IMPORANTE QU EPAR ALOS VECTORI DE ATQUE DE LA MENAZA USSES GUIONE S MEDIOS COMO VIÑETAS DENTRO DE ALS CELDAS"
        prompt = prompt & "Se detecto mediante en red privada en sitio red privada, pero explica los escenarios que consideres necesarios" & vbCrLf & Chr(10)
        prompt = prompt & "SOLO DOS columnas, nombre y amenaza"
        prompt = prompt & " " & vbCrLf & Chr(10) & listaVulnerabilidades & vbCrLf & Chr(10)
          
        ' Copiar el prompt generado al portapapeles
        CopyToClipboard prompt
        
        ' Mostrar un mensaje informativo
        MsgBox "El prompt ha sido copiado al portapapeles.", vbInformation, "Prompt Generado"
    Else
        MsgBox "No se encontraron valores en la selección.", vbExclamation, "Sin Datos"
    End If
End Sub


Sub CYB076_PreparePromptFromSelection_PropuestaRemediacionVuln_OnPrivateNetwork()
    Dim celda As Range
    Dim listaVulnerabilidades As String
    Dim prompt As String
    
    ' Inicializar la variable que almacenará las vulnerabilidades
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
        prompt = "Hola, por favor Redacta como un pentester un párrafo técnico de propuesta de remediación que comience con la frase: -Se recomienda…-. Incluye tantos detalles puntuales par remaicion por jemplo nombre de soluciones que funeriona como poroteccon o controles de sguridad, disoitivos, practicas, menciona de manera más puntual que se pod´ria hacer para que le encargado del sistema, activo, pueda saber como remediari La respueste debe tener SOLO parrafo breve introducción y luego viñetas para lso putnos de propueta de remedaicion Responde para la siguientel lista de vulnerabilidades en FORMATO TABLA DOS COLUMNAS, , SIEMPRE COMIENZAX CON -Se recomienda…- TEXTO AMPLIDO MAS DE 80 palabras, atalmente aplicable a muchos casos, lengauje so framworks"
           prompt = prompt & "Se detecto mediante en red privada en sitio red privada, pero explica los escenarios que consideres necesarios" & vbCrLf & Chr(10)
          prompt = prompt & "Menciona solo soluciones corporativas"
          prompt = prompt & "Solo dos columnas, nombre y propuesat remediacion"
        prompt = prompt & " " & vbCrLf & Chr(10) & listaVulnerabilidades & vbCrLf & Chr(10)
          
        ' Copiar el prompt generado al portapapeles
        CopyToClipboard prompt
        
        ' Mostrar un mensaje informativo
        MsgBox "El prompt ha sido copiado al portapapeles.", vbInformation, "Prompt Generado"
    Else
        MsgBox "No se encontraron valores en la selección.", vbExclamation, "Sin Datos"
    End If
End Sub

Sub CYB0761_PreparePromptFromSelection_TecnicalExplainationVuln_OnPrivateNetwork()
    Dim celda As Range
    Dim listaVulnerabilidades As String
    Dim prompt As String
    
    ' Inicializar la variable que almacenará las vulnerabilidades
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
prompt = "Hola, por favor en una tabla, solo dos columnas: vulnerabilidad y explicación técnica. "
prompt = prompt & "Para cada una de estas vulnerabilidades redacta un párrafo de explicación técnica que contenga un ejemplo y "
prompt = prompt & "una conclusión breve, concisa y convincente desde la perspectiva de pentesting. "
prompt = prompt & "Inicia la explicación con el texto -En un escenario…- [típico / común / poco probable]. "
prompt = prompt & "De ser posible, agrega código de ejemplo para comprender este tipo de vulnerabilidad. "
prompt = prompt & "El código debe ser útil, no seas escaso en detalles. "
prompt = prompt & "NO MENCIONES RECOMENDACIONES. "
prompt = prompt & "Ejemplo: "
prompt = prompt & "-En un escenario… "
prompt = prompt & "Ejemplo: "
prompt = prompt & "Set-Cookie: sessionID=12345; "
prompt = prompt & "String filePath = -/data/- + userInput + -.txt-; "
prompt = prompt & "public fun{} "
prompt = prompt & "Etc…. "
prompt = prompt & "Se considera inseguro o una vulnerabilidad debido a que… "
prompt = prompt & "En conclusión, esta vulnerabilidad es [POTENCIALMENTE EXPLOTABLE] en lo que respecta al código estático. "
prompt = prompt & "QUIERO UNA TABLA CON BUEN FORMATO EN LAS CELDAS, SALTOS DE LÍNEA APROPIADOS. "
prompt = prompt & "MÁS DE 125 CARACTERES. "
prompt = prompt & "NO PONGAS TODO EN UN SOLO PÁRRAFO, USA SALTOS DE LÍNEA DENTRO DE LAS CELDAS DE EXPLICACIÓN PARA QUE SEA LEGIBLE. "
        prompt = prompt & " " & vbCrLf & Chr(10) & listaVulnerabilidades & vbCrLf & Chr(10)
        
        ' Copiar el prompt generado al portapapeles
        CopyToClipboard prompt
        
        ' Mostrar un mensaje informativo
        MsgBox "El prompt ha sido copiado al portapapeles.", vbInformation, "Prompt Generado"
    Else
        MsgBox "No se encontraron valores en la selección.", vbExclamation, "Sin Datos"
    End If
End Sub

Sub CYB0762_PreparePromptFromSelection_CVSSVectorVuln_OnPrivateNetwork()
    Dim celda As Range
    Dim listaVulnerabilidades As String
    Dim prompt As String
    
    ' Inicializar la variable que almacenará las vulnerabilidades
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
        CopyToClipboard prompt
        
        ' Mostrar un mensaje informativo
        MsgBox "El prompt ha sido copiado al portapapeles.", vbInformation, "Prompt Generado"
    Else
        MsgBox "No se encontraron valores en la selección.", vbExclamation, "Sin Datos"
    End If
End Sub

Sub CYB077_PreparePromptFromSelection_DescripcionVuln_FromInternet()
    Dim celda As Range
    Dim listaVulnerabilidades As String
    Dim prompt As String
    
    ' Inicializar la variable que almacenará las vulnerabilidades
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
        prompt = "Redacta un párrafo técnico breve y conciso que describa la vulnerabilidad detectada, comenzando con la frase: -El sistema…-. Explica en qué consiste la debilidad de seguridad de manera técnica. No incluyas escenarios de explotación, ya que eso corresponde a otro campo. No describas cómo se explota, solo en qué consiste el problema. No menciones el nombre exacto de la vulnerabilidad; utiliza expresiones similares. Vulnerabilidades a describi en formato tabla: "
        prompt = prompt & " " & vbCrLf & Chr(10) & listaVulnerabilidades & vbCrLf & Chr(10)
        prompt = prompt & "Contexto de análisis: Análisis de vulnerabilidades de infraestructura a partir conexion a red en desde internet en sitio. Vulnerabilidad detectada mediante escaneo e interacciones desde desde internet en desde internet en sitio."

        
        ' Copiar el prompt generado al portapapeles
        CopyToClipboard prompt
        
        ' Mostrar un mensaje informativo
        MsgBox "El prompt ha sido copiado al portapapeles.", vbInformation, "Prompt Generado"
    Else
        MsgBox "No se encontraron valores en la selección.", vbExclamation, "Sin Datos"
    End If
End Sub



Sub CYB078_PreparePromptFromSelection_AmenazaVuln_FromInternet()
    Dim celda As Range
    Dim listaVulnerabilidades As String
    Dim prompt As String
    
    ' Inicializar la variable que almacenará las vulnerabilidades
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
        prompt = "Hola, por favor considera el siguiente ejemplo: Un atacante podría (obtener, realizar, ejecutar, visualizar, identificar, listar)… Algunos posibles vectores de ataque adicionales asociados con esta amenaza incluyen (pueden ser algunos de los siguientes, pero no necesariamente todos): •  Malware: Un malware diseñado para automatizar intentos de fuerza bruta podría explotar la vulnerabilidad para… •    Usuario malintencionado: Un usuario dentro de la red local con conocimiento de la vulnerabilidad podría aprovecharla para… •    Personal interno: Un empleado con acceso y conocimientos técnicos podría, intencionalmente o por error,… •  Delincuente cibernético: Un atacante externo en busca de vulnerabilidades podría intentar explotar esta debilidad para… Instrucciones adicionales: 1.   Pregunta si el sistema es interno o externo para determinar los vectores de ataque más relevantes, ya que no todos aplican en todos los casos. "
        prompt = prompt & "2.   Redacta una descripción de la amenaza que incluya posibles escenarios de ataque, comenzando con la frase: -Un atacante podría…-. 3.    No es necesario proporcionar un ejemplo para cada vector de ataque. Selecciona el más realista o probable. 4.   En las viñetas de vectores de ataque, menciona la probabilidad de ocurrencia, incluso si es baja. 5.    El contexto es un análisis de vulnerabilidades de infraestructura desde una desde internet, específicamente: -Vulnerabilidad detectada mediante escaneo e interacciones desde desde internet-. Formato de respuesta: •    Responde en una tabla de dos columnas. •    Para cada vulnerabilidad, redacta un párrafo descriptivo en la primera columna (mínimo 75 palabras). •  En la segunda columna, lista los vectores de ataque con viñetas (usando guiones  - ). • No uses HTML, solo texto plano. Ejemplo de estructura: Descripción de la amenaza    Vectores de ataque Un atacante podría explotar esta vulnerabilidad para acceder a información"
        prompt = prompt & "confidencial Esta amenaza es particularmente crítica en sistemas internos donde los controles de seguridad son menos estrictos."
        prompt = prompt & "Un escenario probable incluye...    - Malware: Un malware podría ser utilizado para... (probabilidad media). - Usuario malintencionado: Un empleado con acceso podría... (probabilidad baja). - Delincuente cibernético: Un atacante externo podría... (probabilidad alta). ES MUY IMPORANTE QU EPAR ALOS VECTORI DE ATQUE DE LA MENAZA USSES GUIONE S MEDIOS COMO VIÑETAS DENTRO DE ALS CELDAS"
        prompt = prompt & "Se detecto mediante en desde internet en sitio desde internet, pero explica los escenarios que consideres necesarios" & vbCrLf & Chr(10)
        prompt = prompt & "SOLO DOS columnas, nombre y amenaza"
        prompt = prompt & " " & vbCrLf & Chr(10) & listaVulnerabilidades & vbCrLf & Chr(10)
          
        ' Copiar el prompt generado al portapapeles
        CopyToClipboard prompt
        
        ' Mostrar un mensaje informativo
        MsgBox "El prompt ha sido copiado al portapapeles.", vbInformation, "Prompt Generado"
    Else
        MsgBox "No se encontraron valores en la selección.", vbExclamation, "Sin Datos"
    End If
End Sub


Sub CYB079_PreparePromptFromSelection_PropuestaRemediacionVuln_FromInternet()
    Dim celda As Range
    Dim listaVulnerabilidades As String
    Dim prompt As String
    
    ' Inicializar la variable que almacenará las vulnerabilidades
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
        prompt = "Hola, por favor Redacta como un pentester un párrafo técnico de propuesta de remediación que comience con la frase: -Se recomienda…-. Incluye tantos detalles puntuales par remaicion por jemplo nombre de soluciones que funeriona como poroteccon o controles de sguridad, disoitivos, practicas, menciona de manera más puntual que se pod´ria hacer para que le encargado del sistema, activo, pueda saber como remediari La respueste debe tener SOLO parrafo breve introducción y luego viñetas para lso putnos de propueta de remedaicion Responde para la siguientel lista de vulnerabilidades en FORMATO TABLA DOS COLUMNAS, , SIEMPRE COMIENZAX CON -Se recomienda…- TEXTO AMPLIDO MAS DE 80 palabras, atalmente aplicable a muchos casos, lengauje so framworks"
           prompt = prompt & "Se detecto mediante en desde internet en sitio desde internet, pero explica los escenarios que consideres necesarios" & vbCrLf & Chr(10)
          prompt = prompt & "Menciona solo soluciones corporativas"
          prompt = prompt & "Solo dos columnas, nombre y propuesat remediacion"
        prompt = prompt & " " & vbCrLf & Chr(10) & listaVulnerabilidades & vbCrLf & Chr(10)
          
        ' Copiar el prompt generado al portapapeles
        CopyToClipboard prompt
        
        ' Mostrar un mensaje informativo
        MsgBox "El prompt ha sido copiado al portapapeles.", vbInformation, "Prompt Generado"
    Else
        MsgBox "No se encontraron valores en la selección.", vbExclamation, "Sin Datos"
    End If
End Sub

Sub CYB080_PreparePromptFromSelection_DescripcionVuln_FromCode()
    Dim celda As Range
    Dim listaVulnerabilidades As String
    Dim prompt As String
    
    ' Inicializar la variable que almacenará las vulnerabilidades
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
        prompt = prompt & "Redacta un párrafo técnico breve y conciso que describa la vulnerabilidad detectada en el código fuente, comenzando con la frase: " & vbCrLf
        prompt = prompt & "-El código…-. Explica en qué consiste la debilidad de seguridad de manera técnica. " & vbCrLf
        prompt = prompt & "No incluyas escenarios de explotación, ya que eso corresponde a otro campo. " & vbCrLf
        prompt = prompt & "No describas cómo se explota, solo en qué consiste el problema. " & vbCrLf
        prompt = prompt & "No menciones el nombre exacto de la vulnerabilidad; utiliza expresiones similares." & vbCrLf & vbCrLf
        prompt = prompt & "Vulnerabilidades a describir en formato tabla:" & vbCrLf
        prompt = prompt & listaVulnerabilidades & vbCrLf & vbCrLf
        prompt = prompt & "Contexto de análisis: " & vbCrLf
        prompt = prompt & "Análisis de código fuente mediante herramientas de análisis estático (SAST). " & vbCrLf
        prompt = prompt & "Vulnerabilidad identificada a partir del análisis del código sin ejecutar la aplicación." & vbCrLf

        ' Copiar el prompt generado al portapapeles
        CopyToClipboard prompt
        
        ' Mostrar un mensaje informativo
        MsgBox "El prompt ha sido copiado al portapapeles.", vbInformation, "Prompt Generado"
    Else
        MsgBox "No se encontraron valores en la selección.", vbExclamation, "Sin Datos"
    End If
End Sub

Sub CYB081_PreparePromptFromSelection_AmenazaVuln_FromCode()
    Dim celda As Range
    Dim listaVulnerabilidades As String
    Dim prompt As String
    
    ' Inicializar la variable que almacenará las vulnerabilidades
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
        prompt = prompt & "Un atacante podría (inyectar, manipular, filtrar, exponer, escalar privilegios)…" & vbCrLf
        prompt = prompt & "Algunos posibles vectores de ataque adicionales asociados con esta amenaza incluyen (pueden ser algunos de los siguientes, pero no necesariamente todos):" & vbCrLf
        prompt = prompt & "•  Inyección de código: Una entrada no validada en el código fuente podría permitir inyección de comandos..." & vbCrLf
        prompt = prompt & "•  Exposición de información sensible: Una mala gestión de credenciales en el código podría revelar secretos..." & vbCrLf
        prompt = prompt & "•  Elevación de privilegios: Una función mal diseñada podría permitir a un usuario ejecutar acciones con más permisos de los necesarios..." & vbCrLf
        prompt = prompt & "•  Manipulación de datos: Un atacante podría modificar parámetros dentro del código para alterar la lógica de la aplicación..." & vbCrLf
        prompt = prompt & vbCrLf
        prompt = prompt & "Instrucciones adicionales:" & vbCrLf
        prompt = prompt & "1. Pregunta si el código pertenece a una aplicación interna o externa para determinar los vectores de ataque más relevantes." & vbCrLf
        prompt = prompt & "2. Redacta una descripción de la amenaza que incluya posibles escenarios de ataque, comenzando con la frase: -Un atacante podría…-." & vbCrLf
        prompt = prompt & "3. No es necesario proporcionar un ejemplo para cada vector de ataque. Selecciona el más realista o probable." & vbCrLf
        prompt = prompt & "4. En las viñetas de vectores de ataque, menciona la probabilidad de ocurrencia, incluso si es baja." & vbCrLf
        prompt = prompt & "5. El contexto es un análisis de código fuente mediante herramientas de análisis estático (SAST), sin ejecutar la aplicación." & vbCrLf
        prompt = prompt & vbCrLf
        prompt = prompt & "Formato de respuesta:" & vbCrLf
        prompt = prompt & "• Responde en una tabla de dos columnas." & vbCrLf
        prompt = prompt & "• Para cada vulnerabilidad, redacta un párrafo descriptivo en la primera columna (mínimo 75 palabras)." & vbCrLf
        prompt = prompt & "• En la segunda columna, lista los vectores de ataque con viñetas (usando guiones - )." & vbCrLf
        prompt = prompt & "• No uses HTML, solo texto plano." & vbCrLf
        prompt = prompt & vbCrLf
        prompt = prompt & "Ejemplo de estructura:" & vbCrLf
        prompt = prompt & "Descripción de la amenaza    Vectores de ataque" & vbCrLf
        prompt = prompt & "Un atacante podría explotar esta vulnerabilidad en el código para ejecutar comandos arbitrarios..." & vbCrLf
        prompt = prompt & "    - Inyección de código: Un usuario malintencionado podría insertar código malicioso... (probabilidad media)." & vbCrLf
        prompt = prompt & "    - Exposición de información: Credenciales en el código podrían filtrarse... (probabilidad alta)." & vbCrLf
        prompt = prompt & vbCrLf
        prompt = prompt & "ES MUY IMPORTANTE QUE PARA LOS VECTORES DE ATAQUE DE LA AMENAZA USES GUIONES MEDIOS COMO VIÑETAS DENTRO DE LAS CELDAS." & vbCrLf
        prompt = prompt & "Análisis realizado mediante herramientas SAST en código fuente estático sin ejecución." & vbCrLf
        prompt = prompt & "SOLO DOS COLUMNAS: NOMBRE Y AMENAZA." & vbCrLf
        prompt = prompt & vbCrLf & listaVulnerabilidades & vbCrLf & vbCrLf
          
        ' Copiar el prompt generado al portapapeles
        CopyToClipboard prompt
        
        ' Mostrar un mensaje informativo
        MsgBox "El prompt ha sido copiado al portapapeles.", vbInformation, "Prompt Generado"
    Else
        MsgBox "No se encontraron valores en la selección.", vbExclamation, "Sin Datos"
    End If
End Sub


Sub CYB082_PreparePromptFromSelection_PropuestaRemediacionVuln_FromCode()
    Dim celda As Range
    Dim listaVulnerabilidades As String
    Dim prompt As String
    
    ' Inicializar la variable que almacenará las vulnerabilidades
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
        prompt = "Hola, por favor redacta como un pentester un párrafo técnico de propuesta de remediación que comience con la frase: -Se recomienda…-."
        prompt = prompt & " Incluye tantos detalles puntuales como sea posible, mencionando soluciones específicas, controles de seguridad, dispositivos y prácticas recomendadas."
        prompt = prompt & " Proporciona información clara para que el encargado del sistema o activo sepa exactamente cómo remediarlo."
        prompt = prompt & " La respuesta debe contener un párrafo breve de introducción seguido de viñetas con los puntos de la propuesta de remediación."
        prompt = prompt & " Formato de respuesta: una tabla de dos columnas."
        prompt = prompt & " Siempre comienza con -Se recomienda…-."
        prompt = prompt & " El texto debe ser amplio, con más de 80 palabras, aplicable a múltiples casos y en lenguaje técnico adecuado."
        prompt = prompt & " Se detectó mediante análisis desde internet en el sitio, pero explica los escenarios relevantes."
        prompt = prompt & " Menciona solo soluciones corporativas."
        prompt = prompt & " Solo dos columnas: nombre y propuesta de remediación."
        prompt = prompt & " " & vbCrLf & Chr(10) & listaVulnerabilidades & vbCrLf & Chr(10)
        
        ' Copiar el prompt generado al portapapeles
        CopyToClipboard prompt
        
        ' Mostrar un mensaje informativo
        MsgBox "El prompt ha sido copiado al portapapeles.", vbInformation, "Prompt Generado"
    Else
        MsgBox "No se encontraron valores en la selección.", vbExclamation, "Sin Datos"
    End If
End Sub


Sub CYB083_Verify_CVSS4_0_Vector()
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
    
    ' Verificar si se encontró un vector CVSS
    If cvssString <> "" Then
        ' Construir la URL
        url = "https://www.first.org/cvss/calculator/4.0#" & cvssString
        
        ' Copiar la URL al portapapeles
        CopyToClipboard url
        
        ' Mostrar un mensaje con la URL copiada
        MsgBox "La URL ha sido copiada al portapapeles: " & vbCrLf & url, vbInformation, "URL Generada"
        
        ' Abrir la URL en el navegador
        ThisWorkbook.FollowHyperlink url
    Else
        MsgBox "No se encontraron valores en la selección.", vbExclamation, "Sin Datos"
    End If
End Sub


Sub CopyToClipboard(text As String)
    ' Copiar texto al portapapeles usando Microsoft Forms
    Dim MSForms_DataObject As Object
    Set MSForms_DataObject = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    MSForms_DataObject.SetText text
    MSForms_DataObject.PutInClipboard
    Set MSForms_DataObject = Nothing
End Sub


