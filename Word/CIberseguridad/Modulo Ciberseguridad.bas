Attribute VB_Name = "NewMacros"
Sub FormatearTablaVulnerabilidadesAvanzado()

    Dim tbl As Table
    Dim i As Long ' Índice para filas
    Dim j As Long ' Índice para columnas (al escanear encabezado)
    
    Dim cellText As String         ' Para el texto de las celdas de datos
    Dim headerCellText As String   ' Para el texto de las celdas del encabezado
    Dim estadoCellText As String   ' Para el texto de la celda de estado
    
    ' Índices de las columnas (0 si no se encuentran)
    Dim severityColIndex As Integer
    Dim vulnerabilityNameColIndex As Integer
    Dim estadoColIndex As Integer
    
    severityColIndex = 0
    vulnerabilityNameColIndex = 0
    estadoColIndex = 0
    
    ' --- Variables de Color (RGB) ---
    Dim HEADER_BLUE As Long
    Dim CRITICA_PURPLE As Long
    Dim ALTA_RED As Long
    Dim MEDIA_YELLOW As Long
    Dim BAJA_GREEN As Long
    Dim ZEBRA_LIGHT_GRAY As Long
    
    ' Colores para el texto de la columna "Estado"
    Dim ESTADO_SIN_REMEDIAR_COLOR As Long
    Dim ESTADO_REMEDIADA_COLOR As Long
    Dim ESTADO_NUEVA_COLOR As Long

    ' Asignar valores a las variables de color
    HEADER_BLUE = RGB(0, 112, 192)      ' #0070C0
    CRITICA_PURPLE = RGB(112, 48, 160)   ' #7030A0
    ALTA_RED = RGB(255, 0, 0)           ' #FF0000
    MEDIA_YELLOW = RGB(255, 255, 0)     ' #FFFF00
    BAJA_GREEN = RGB(0, 176, 80)        ' #00B050 (Verde)
    ZEBRA_LIGHT_GRAY = RGB(242, 242, 242) ' Gris muy claro para zebra
    
    ESTADO_SIN_REMEDIAR_COLOR = RGB(255, 0, 0)   ' Rojo para "Sin remediar"
    ESTADO_REMEDIADA_COLOR = RGB(0, 176, 80)       ' Verde oscuro para "Remediada"
    ESTADO_NUEVA_COLOR = RGB(255, 192, 0)      ' Naranja/Ámbar para "Nueva detección"


    ' Verificar si hay una tabla seleccionada
    If Selection.Information(wdWithInTable) = False Then
        MsgBox "Por favor, selecciona una tabla primero.", vbExclamation, "Error"
        Exit Sub
    End If
    
    Set tbl = Selection.Tables(1)
    
    ' --- 1. Encontrar Índices de Columnas desde el Encabezado ---
    If tbl.Rows.Count >= 1 Then
        For j = 1 To tbl.Rows(1).Cells.Count
            Dim rawHeaderText As String
            rawHeaderText = tbl.Rows(1).Cells(j).Range.Text
            
            ' Limpiar texto del encabezado
            headerCellText = Replace(rawHeaderText, Chr(7), "")
            headerCellText = Replace(headerCellText, vbCrLf, "")
            headerCellText = Replace(headerCellText, vbCr, "")
            headerCellText = Replace(headerCellText, vbLf, "")
            headerCellText = Trim(headerCellText)
            
            Select Case LCase(headerCellText)
                Case "severidad"
                    severityColIndex = j
                Case "nombre de vulnerabilidad"
                    vulnerabilityNameColIndex = j
                Case "estado"
                    estadoColIndex = j
            End Select
        Next j
    End If

    ' --- 2. Formato del Encabezado ---
    If tbl.Rows.Count >= 1 Then
        With tbl.Rows(1)
            .Shading.BackgroundPatternColor = HEADER_BLUE
            .Range.Font.ColorIndex = wdWhite
            .Range.Font.Bold = True
            .Cells.VerticalAlignment = wdCellAlignVerticalCenter ' Centrado Vertical
            Dim cl As Cell
            For Each cl In .Cells
                cl.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter ' Centrado Horizontal
            Next cl
        End With
    End If

    ' --- 3. Ajustar Ancho de Columna "Nombre de vulnerabilidad" ---
    If vulnerabilityNameColIndex > 0 Then
        tbl.Columns(vulnerabilityNameColIndex).SetWidth CentimetersToPoints(9), RulerStyle:=wdAdjustNone
    End If

    ' --- 4. Formato de Filas de Datos (Zebra, Severidad y Estado) ---
    If tbl.Rows.Count >= 2 Then
        For i = 2 To tbl.Rows.Count
            
            ' Aplicar formato Zebra a toda la fila primero
            If (i Mod 2) = 0 Then
                tbl.Rows(i).Shading.BackgroundPatternColor = ZEBRA_LIGHT_GRAY
            Else
                tbl.Rows(i).Shading.BackgroundPatternColor = wdColorWhite
            End If
            
            ' Color de fuente y estilo por defecto para la fila de datos
            tbl.Rows(i).Range.Font.ColorIndex = wdBlack
            tbl.Rows(i).Range.Font.Bold = False

            ' Formato condicional para la celda de "Severidad"
            If severityColIndex > 0 Then
                If tbl.Rows(i).Cells.Count >= severityColIndex Then
                    With tbl.Cell(i, severityColIndex)
                        If .Range.Characters.Count > 0 Then
                             Dim tempSeverityText As String
                             tempSeverityText = .Range.Text
                             tempSeverityText = Replace(tempSeverityText, Chr(7), "")
                             tempSeverityText = Replace(tempSeverityText, vbCrLf, "")
                             tempSeverityText = Replace(tempSeverityText, vbCr, "")
                             tempSeverityText = Replace(tempSeverityText, vbLf, "")
                             cellText = Trim(tempSeverityText)
                        Else
                            cellText = ""
                        End If
                        
                        Select Case UCase(cellText)
                            Case "CRÍTICA"
                                .Shading.BackgroundPatternColor = CRITICA_PURPLE
                                .Range.Font.ColorIndex = wdWhite
                            Case "ALTA"
                                .Shading.BackgroundPatternColor = ALTA_RED
                                .Range.Font.ColorIndex = wdWhite
                            Case "MEDIA"
                                .Shading.BackgroundPatternColor = MEDIA_YELLOW
                                .Range.Font.ColorIndex = wdBlack
                            Case "BAJA"
                                .Shading.BackgroundPatternColor = BAJA_GREEN
                                .Range.Font.ColorIndex = wdWhite
                            Case Else
                                If (i Mod 2) = 0 Then
                                    .Shading.BackgroundPatternColor = ZEBRA_LIGHT_GRAY
                                Else
                                    .Shading.BackgroundPatternColor = wdColorWhite
                                End If
                                .Range.Font.ColorIndex = wdBlack
                        End Select
                        .Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                        .VerticalAlignment = wdCellAlignVerticalCenter
                    End With
                End If
            End If
            
            ' Formato condicional para la celda de "Estado"
            If estadoColIndex > 0 Then
                If tbl.Rows(i).Cells.Count >= estadoColIndex Then
                    With tbl.Cell(i, estadoColIndex)
                        If .Range.Characters.Count > 0 Then
                             Dim tempEstadoText As String
                             tempEstadoText = .Range.Text
                             tempEstadoText = Replace(tempEstadoText, Chr(7), "")
                             tempEstadoText = Replace(tempEstadoText, vbCrLf, "")
                             tempEstadoText = Replace(tempEstadoText, vbCr, "")
                             tempEstadoText = Replace(tempEstadoText, vbLf, "")
                             estadoCellText = Trim(tempEstadoText)
                        Else
                            estadoCellText = ""
                        End If
                        
                        ' Guardar color de fuente original por si no hay match
                        Dim originalFontColor As Long
                        originalFontColor = .Range.Font.ColorIndex

                        ' Aplicar color de fuente según el texto (buscando palabras clave)
                        If InStr(1, UCase(estadoCellText), "SIN REMEDIAR") > 0 Then
                            .Range.Font.color = ESTADO_SIN_REMEDIAR_COLOR
                        ElseIf InStr(1, UCase(estadoCellText), "REMEDIADA") > 0 Then
                            .Range.Font.color = ESTADO_REMEDIADA_COLOR
                        ElseIf InStr(1, UCase(estadoCellText), "NUEVA DETECCIÓN") > 0 Then
                            .Range.Font.color = ESTADO_NUEVA_COLOR
                        Else
                            ' Si no coincide, la fuente mantiene el color por defecto de la fila (negro)
                            .Range.Font.ColorIndex = wdBlack
                        End If
                        ' Aquí puedes añadir alineación para la celda de Estado si es necesario
                        ' .Range.ParagraphFormat.Alignment = wdAlignParagraphLeft ' o Center
                        .VerticalAlignment = wdCellAlignVerticalCenter
                    End With
                End If
            End If
        Next i
    End If
    
    ' --- 5. Aplicar Bordes ---
    With tbl.Borders
        .Enable = True
        .InsideLineStyle = wdLineStyleSingle
        .OutsideLineStyle = wdLineStyleSingle
        .InsideColor = wdColorAutomatic
        .OutsideColor = wdColorAutomatic
    End With

    Selection.Collapse Direction:=wdCollapseEnd
    
    Dim feedbackMsg As String
    feedbackMsg = "Formato de tabla de vulnerabilidades aplicado."
    If severityColIndex = 0 Then
        feedbackMsg = feedbackMsg & vbCrLf & "ADVERTENCIA: No se encontró la columna 'Severidad'."
    End If
    If vulnerabilityNameColIndex = 0 Then
        feedbackMsg = feedbackMsg & vbCrLf & "ADVERTENCIA: No se encontró la columna 'Nombre de vulnerabilidad'."
    End If
    If estadoColIndex = 0 Then
        feedbackMsg = feedbackMsg & vbCrLf & "ADVERTENCIA: No se encontró la columna 'Estado'."
    End If
    
    MsgBox feedbackMsg, vbInformation, "Macro Finalizada"

End Sub
Sub NegritaPalabrasClave_Robusta_MultiArray_Corregido()

    Dim palabrasClaveParte1 As Variant
    Dim palabrasClaveParte2 As Variant
    Dim palabrasClaveParte3 As Variant
    Dim todasLasPalabrasClaveList As Object
    Dim palabrasClaveOrdenadas As Variant
    Dim palabra As Variant
    Dim i As Long
    Dim temp As String
    Dim swapped As Boolean
    Dim rangoAProcesar As Range
    Dim findObj As Find

    ' Definición de palabras clave en fragmentos
    palabrasClaveParte1 = Array( _
        "XSS", "Stored XSS", "DOM-based XSS", "versiones obsoletas", _
        "Session Hijacking", "Phishing", "CSP", _
        "Validación de entradas", "Sanitización", "Interceptación", "interceptar", _
        "TLS 1.0", "Protocolo débil", "Man-in-the-Middle" _
    )

    palabrasClaveParte2 = Array( _
        "facilita ataques", "Explotación", "TLS 1.1", "controles de restricción adecuados", _
        "Downgrade Attack", "Tráfico TLS", "Sweet32", _
        "Wireshark", "Tshark", "tcpdump", _
        "Gestión de vulnerabilidades", "Seguridad de aplicaciones web", "Divulgación de información" _
    )

    palabrasClaveParte3 = Array( _
        "Fuga de información", "Hardening", "HTTP Headers", _
        "Autodiscover", "X-Frame-Options", "HttpOnly", _
        "SSL Stripping", "Componentes vulnerables", _
        "OWASP DependencyCheck", "Subresource Integrity", "Clickjacking", _
        "Cookies HttpOnly", "Cookies Secure", "Fingerprinting", _
        "CVE", _
        "Metodología de pentesting", "Seguridad en la nube", "Acceso no autorizado", _
        "Credenciales", "Confidencialidad", "Remediación" _
    )
    
    Dim palabrasClaveReporteSMB As Variant

palabrasClaveReporteSMB = Array( _
    "firma SMB", "validación criptográfica", "interceptado", "alterado", _
    "vulnerabilidad", "MITM", "ataque Man-in-the-Middle", "manipular", _
    "red interna", "remediar", "habilitada", "autenticidad", "integridad", _
    "políticas de grupo", "GPO", "actualizar sistemas operativos", "parches de seguridad", _
    "versiones obsoletas de SMB", "SMBv1", "SMBv2", "SMBv3" _
)
    ' Crear ArrayList para juntar todas las palabras
    Set todasLasPalabrasClaveList = CreateObject("System.Collections.ArrayList")

    ' Añadir palabras de cada array a la lista
    For Each palabra In palabrasClaveParte1
        If Len(Trim(palabra)) > 0 Then todasLasPalabrasClaveList.Add Trim(palabra)
    Next
    For Each palabra In palabrasClaveParte2
        If Len(Trim(palabra)) > 0 Then todasLasPalabrasClaveList.Add Trim(palabra)
    Next
    For Each palabra In palabrasClaveParte3
        If Len(Trim(palabra)) > 0 Then todasLasPalabrasClaveList.Add Trim(palabra)
    Next
    
    
      For Each palabra In palabrasClaveReporteSMB
        If Len(Trim(palabra)) > 0 Then todasLasPalabrasClaveList.Add Trim(palabra)
    Next

    ' Ordenar por longitud descendente para evitar conflictos en palabras que son substrings de otras
    Do
        swapped = False
        For i = 0 To todasLasPalabrasClaveList.Count - 2
            If Len(todasLasPalabrasClaveList(i)) < Len(todasLasPalabrasClaveList(i + 1)) Then
                temp = todasLasPalabrasClaveList(i)
                todasLasPalabrasClaveList(i) = todasLasPalabrasClaveList(i + 1)
                todasLasPalabrasClaveList(i + 1) = temp
                swapped = True
            End If
        Next i
    Loop While swapped

    palabrasClaveOrdenadas = todasLasPalabrasClaveList.ToArray()

    ' Definir el rango a procesar (selección)
    If Selection.Type = wdSelectionIP Or Selection.Type = wdNoSelection Then
        MsgBox "Por favor, seleccione el texto donde desea aplicar la negrita.", vbExclamation
        Exit Sub
    End If
    Set rangoAProcesar = Selection.Range

    Application.ScreenUpdating = False

    ' Buscar y aplicar negrita para cada palabra clave
    For Each palabra In palabrasClaveOrdenadas
        Set findObj = rangoAProcesar.Duplicate.Find
        With findObj
            .ClearFormatting
            .Text = palabra
            .Replacement.ClearFormatting
            .Replacement.Text = "^&" ' Mantener el texto encontrado
            .Replacement.Font.Bold = True
            .Forward = True
            .Wrap = wdFindStop ' Importante usar wdFindStop en rangos
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False ' Cambiar a True si quieres solo palabras completas
            .MatchWildcards = False
            .Execute Replace:=wdReplaceAll
        End With
    Next palabra

    Application.ScreenUpdating = True

    MsgBox "Palabras clave en negrita aplicadas.", vbInformation

End Sub


Sub AjustarFormatoColumnasTablaVulnes()
    Dim tbl As Table
    Dim colWidths(1 To 4) As Double

    ' Definir anchos en cm
    colWidths(1) = 2
    colWidths(2) = 7
    colWidths(3) = 5
    colWidths(4) = 4

    ' Verifica que haya una tabla seleccionada
    If Selection.Information(wdWithInTable) Then
        Set tbl = Selection.Tables(1)

        ' Desactivar ajuste automático de columnas
        tbl.AllowAutoFit = False

        ' Centrar la tabla en la página
        tbl.Rows.Alignment = wdAlignRowCenter

        ' Ajustar fuente de encabezado
        tbl.Rows(1).Range.Font.Size = 11

        ' Ajustar fuente del resto de filas
        Dim r As Integer
        For r = 2 To tbl.Rows.Count
            tbl.Rows(r).Range.Font.Size = 10
        Next r

        ' Ajustar anchos de columnas (fijos)
        Dim i As Integer
        For i = 1 To 4
            If i <= tbl.Columns.Count Then
                With tbl.Columns(i)
                    .PreferredWidthType = wdPreferredWidthPoints
                    .PreferredWidth = CentimetersToPoints(colWidths(i))
                    .Width = CentimetersToPoints(colWidths(i))
                End With
            End If
        Next i

    Else
        MsgBox "Por favor selecciona una tabla primero.", vbExclamation
    End If
End Sub











Sub NegritaPalabrasClave()
    Dim palabrasTexto As String
    Dim palabrasClave As Variant
    Dim palabra As Variant
    Dim rango As Range
    

palabrasTexto = "Cross-Site Scripting (XSS), XSS Reflected, Reflected XSS, XSS Persistent, Stored XSS, DOM-based XSS, Inyección, Scripts maliciosos, Malicious Scripts, Ejecución de acciones en nombre del usuario, Robo de cookies de sesión, Session Hijacking, Redirección a sitios fraudulentos, Phishing, Content Security Policy (CSP), Validación, Sanitización, Manipulación de sesiones de usuario, Formulario malicioso, TLS 1.0, Protocolo débil, Algoritmos de cifrado, Interceptación, Hombre en el medio, Escucha pasiva, Malware, Explotación,TLS 1.1, downgrade, Interceptar tráfico cifrado, Atacante, Wireshark, Tshark, tcpdump, Tráfico TLS, Sweet32, Colisiones, CBC, (Cipher Block Chaining)"


    palabrasClave = Split(palabrasTexto, ",")

    Set rango = Selection.Range
    
    For Each palabra In palabrasClave
        With rango.Find
            .ClearFormatting
            .Text = Trim(palabra)
            .Replacement.ClearFormatting
            .Replacement.Font.Bold = True
            .Forward = True
            .Wrap = wdFindStop
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .Execute Replace:=wdReplaceAll
        End With
    Next palabra
    
    MsgBox "Palabras clave en negrita aplicadas.", vbInformation
End Sub







Sub FormatearCodigoHTMLSeleccionado()
    Const colorHTML As Long = &H66CC       ' Azul suave para HTML (ej. etiquetas)
    Const colorCSS As Long = &H228B22      ' Verde más claro para CSS (ej. propiedades)
    Const colorJS As Long = &HFF6347       ' Tomate suave para JS (ej. palabras reservadas)
    Const colorValores As Long = &HFF8C00   ' Naranja suave para valores (ej. "auto", "#fff", "true")
    Const colorComentarios As Long = &H808080 ' Gris para comentarios (Simplificado)
    Const colorStrings As Long = &HFF4500   ' Naranja oscuro para cadenas (Simplificado)
    Const colorPorDefecto As Long = &H0      ' Negro para el resto del texto

    Dim htmlKeywords As Variant
    Dim cssKeywords As Variant
    Dim jsKeywords As Variant
    Dim valueKeywords As Variant ' Palabras clave comunes para valores
    Dim keyword As Variant
    Dim word As Range
    Dim selRange As Range

    ' Listas de palabras clave (puedes expandirlas)
    htmlKeywords = Array("html", "body", "head", "title", "meta", "link", "style", "script", "div", "span", "p", "a", "img", "ul", "ol", "li", "table", "tr", "td", "th", "form", "input", "button", "select", "option", "textarea", "label", "h1", "h2", "h3", "h4", "h5", "h6", "header", "footer", "nav", "article", "section", "aside", "main", "figure", "figcaption", "br", "hr", "DOCTYPE")
    cssKeywords = Array("color", "background-color", "background", "font-size", "font-family", "font-weight", "font-style", "text-align", "text-decoration", "padding", "margin", "border", "width", "height", "display", "position", "top", "right", "bottom", "left", "float", "clear", "overflow", "z-index", "opacity", "border-radius", "box-shadow", "transition", "transform", "@media", "@keyframes", "import", "selector", "content", "cursor", "visibility", "list-style")
    jsKeywords = Array("function", "var", "let", "const", "if", "else", "for", "while", "do", "switch", "case", "break", "continue", "return", "try", "catch", "finally", "throw", "new", "this", "class", "extends", "super", "import", "export", "async", "await", "yield", "document", "window", "alert", "console", "log", "error", "warn", "info", "getElementById", "getElementsByTagName", "getElementsByClassName", "querySelector", "querySelectorAll", "addEventListener", "removeEventListener", "setTimeout", "setInterval", "clearTimeout", "clearInterval", "JSON", "parse", "stringify", "Math", "Date", "Array", "Object", "String", "Number", "Boolean")
    valueKeywords = Array("true", "false", "null", "undefined", "auto", "inherit", "initial", "unset", "none", "block", "inline", "inline-block", "flex", "grid", "absolute", "relative", "fixed", "static", "solid", "dotted", "dashed", "double", "hidden", "visible", "bold", "italic", "normal") ' Añadidos valores comunes

    ' ----- Lógica Principal -----
    ' 1. Verificar si hay texto seleccionado
    If Selection.Type = wdSelectionIP Or Selection.Type = wdSelectionNone Then
        MsgBox "Por favor, primero seleccione el bloque de código HTML/CSS/JS que desea formatear.", vbInformation, "Selección Requerida"
        Exit Sub
    End If

    Set selRange = Selection.Range

    ' 2. Restablecer el color de toda la selección a negro (o el color por defecto)
    ' Esto asegura que solo las palabras clave identificadas cambien de color.
    selRange.Font.color = colorPorDefecto

    ' 3. Iterar sobre cada "palabra" en la selección
    ' Usamos Trim() para quitar espacios al inicio/final que a veces incluye .Words
    ' Usamos LCase() para hacer la comparación insensible a mayúsculas/minúsculas
    For Each word In selRange.Words
        Dim wordText As String
        wordText = LCase(Trim(word.Text))
        
        ' Optimización: Si la palabra está vacía después de Trim, saltarla.
        If Len(wordText) = 0 Then GoTo NextWord

        ' Bandera para saber si ya se coloreó la palabra
        Dim colored As Boolean
        colored = False

        ' Comprobar HTML
        For Each keyword In htmlKeywords
            If wordText = keyword Then
                word.Font.color = colorHTML
                colored = True
                Exit For ' Salir del bucle de HTML si se encuentra coincidencia
            End If
        Next keyword
        If colored Then GoTo NextWord ' Ir a la siguiente palabra si ya se coloreó

        ' Comprobar CSS (solo si no es HTML)
        For Each keyword In cssKeywords
            If wordText = keyword Then
                word.Font.color = colorCSS
                colored = True
                Exit For
            End If
        Next keyword
        If colored Then GoTo NextWord

        ' Comprobar JS (solo si no es HTML ni CSS)
        For Each keyword In jsKeywords
            If wordText = keyword Then
                word.Font.color = colorJS
                colored = True
                Exit For
            End If
        Next keyword
        If colored Then GoTo NextWord

        ' Comprobar Valores comunes (solo si no es ninguna de las anteriores)
        For Each keyword In valueKeywords
            If wordText = keyword Then
                word.Font.color = colorValores
                colored = True
                Exit For
            End If
        Next keyword
        ' No necesita GoTo aquí ya que es la última comprobación de palabras clave

' Etiqueta para saltar a la siguiente iteración del bucle principal
NextWord:
    Next word

    ' ----- Limpieza (Buenas prácticas) -----
    Set word = Nothing
    Set selRange = Nothing

    ' Mensaje Opcional de finalización
    ' MsgBox "Formato de sintaxis aplicado a la selección.", vbInformation

End Sub

Sub FormatearComoCodigoXML()
    ' Declaración de variables
    Dim selText As String
    Dim oTable As Table
    Dim oCell As Cell
    Dim oDoc As Document
    Dim itemRange As Range ' Rango para aplicar formato (reusado)
    Dim arrKeywords As Variant
    Dim arrSymbols As Variant
    Dim i As Long
    Dim strItem As String
    Dim originalRange As Range

    ' --- Obtener documento actual y selección ---
    Set oDoc = ActiveDocument
    
    ' Verificar si el documento está protegido
    If oDoc.ReadOnlyRecommended Or oDoc.ProtectionType <> wdNoProtection Then
        MsgBox "El documento está protegido o es de solo lectura y no se puede modificar.", vbExclamation, "Documento Protegido"
        Exit Sub
    End If
    
    ' Verificar si hay texto seleccionado
    If Selection.Type = wdSelectionIP Or Len(Selection.Text) <= 1 Then
        MsgBox "Por favor, seleccione el texto XML que desea formatear.", vbInformation, "Selección Vacía"
        Exit Sub
    End If

    ' Guardar el texto seleccionado y el rango original
    selText = Selection.Text
    Set originalRange = Selection.Range

    ' --- Definir palabras clave XML (etiquetas, atributos, valores comunes) ---
    ' Esta lista puede ser extendida según sea necesario.
    ' Ejemplo de palabras clave (incluyendo las de tu ejemplo):
    arrKeywords = Array("domain-config", "cleartextTrafficPermitted", "domain", "includeSubdomains", _
                        "xml", "version", "encoding", "xs:schema", "xmlns:xs", "xs:element", "xs:complexType", _
                        "xs:sequence", "xs:attribute", "name", "type", "minOccurs", "maxOccurs", "use", "required", "optional", _
                        "configuration", "appSettings", "connectionStrings", "system.web", "compilation", "authentication", _
                        "authorization", "add", "remove", "key", "value", "mode", "debug", "targetFramework", "httpRuntime", _
                        "true", "false", "item", "property", "class", "method", "parameter")

    ' --- Definir símbolos XML (elementos estructurales) ---
    ' El orden es importante: los símbolos más largos primero para evitar coincidencias parciales (ej: "</" antes que "<").
    arrSymbols = Array("</", "/>", "<", ">", "=", """") ' Incluye comillas dobles

    ' --- Crear la tabla y reemplazar la selección actual ---
    ' Al pasar originalRange (que es Selection.Range), la selección se reemplaza por la tabla.
    ' CORRECCIÓN AQUÍ: NumCols -> NumColumns
    Set oTable = oDoc.Tables.Add(Range:=originalRange, NumRows:=1, NumColumns:=1)
    Set oCell = oTable.Cell(1, 1)

    ' --- Insertar el texto original en la celda y establecer fuente base ---
    oCell.Range.Text = selText
    oCell.Range.Font.Name = "Consolas" ' Fuente monoespaciada recomendada para código
    oCell.Range.Font.Size = 10         ' Tamaño de fuente común para código

    ' --- Aplicar fondo gris a la celda ---
    oCell.Shading.BackgroundPatternColor = RGB(240, 240, 240) ' Gris claro (ej: #F0F0F0)

    ' --- PASO 1: Resaltar valores de atributos (texto entre comillas dobles) ---
    Set itemRange = oCell.Range ' Trabajar dentro del rango de la celda
    With itemRange.Find
        .ClearFormatting
        .Text = """*""" ' Encuentra texto encerrado entre comillas dobles (ej: "valor")
        .MatchWildcards = True ' Habilitar comodines para el asterisco
        .Replacement.ClearFormatting
        .Replacement.Font.color = RGB(0, 128, 0)  ' Verde oscuro para los valores de strings
        .Replacement.Font.Name = "Consolas"       ' Mantener la fuente monoespaciada
        .Replacement.Font.Size = 10               ' Mantener el tamaño de fuente
        .Wrap = wdFindStop                        ' No buscar fuera del rango de la celda
        .Execute Replace:=wdReplaceAll            ' Reemplazar todas las ocurrencias
    End With

    ' --- PASO 2: Formatear palabras clave XML ---
    Set itemRange = oCell.Range ' Resetear el rango a la celda completa para la búsqueda
    For i = LBound(arrKeywords) To UBound(arrKeywords)
        strItem = arrKeywords(i)
        With itemRange.Find
            .ClearFormatting
            .Text = strItem
            .Replacement.ClearFormatting
            .Replacement.Font.Bold = True
            .Replacement.Font.color = RGB(0, 0, 255)   ' Azul para palabras clave
            .Replacement.Font.Name = "Consolas"        ' Mantener fuente
            .Replacement.Font.Size = 10                ' Mantener tamaño
            .MatchCase = True        ' XML es sensible a mayúsculas/minúsculas para etiquetas y atributos
            .MatchWholeWord = True   ' Solo palabras completas
            .MatchWildcards = False
            .Wrap = wdFindStop
            .Execute Replace:=wdReplaceAll
        End With
    Next i

    ' --- PASO 3: Formatear símbolos XML ---
    ' Esto se hace después para que los símbolos (como las comillas) tengan su propio color,
    ' incluso si eran parte del resaltado de "valores de atributos".
    Set itemRange = oCell.Range ' Resetear el rango a la celda completa para la búsqueda
    For i = LBound(arrSymbols) To UBound(arrSymbols)
        strItem = arrSymbols(i)
        With itemRange.Find
            .ClearFormatting
            .Text = strItem
            .Replacement.ClearFormatting
            .Replacement.Font.Bold = False ' Los símbolos generalmente no van en negrita
            .Replacement.Font.color = RGB(165, 42, 42)    ' Rojo oscuro/marrón (similar a "Brown" en HTML) para símbolos
            .Replacement.Font.Name = "Consolas"         ' Mantener fuente
            .Replacement.Font.Size = 10                 ' Mantener tamaño
            .MatchCase = True         ' Los símbolos son sensibles a mayúsculas/minúsculas
            .MatchWholeWord = False   ' Los símbolos pueden estar pegados a otros caracteres (ej: <tag>)
            .MatchWildcards = False
            .Wrap = wdFindStop
            .Execute Replace:=wdReplaceAll
        End With
    Next i

    ' --- Mover el cursor después de la tabla ---
    Dim endRange As Range
    Set endRange = oDoc.Range(Start:=oTable.Range.End, End:=oTable.Range.End)

    ' Verificar si la tabla es el último elemento en el cuerpo del documento
    If oTable.Range.End >= oDoc.Content.End - 1 Then ' -1 para la marca de párrafo final del documento
        ' Si es así, añadir un nuevo párrafo después de la tabla para que el cursor se ubique allí
        oDoc.Content.InsertParagraphAfter
        ' Mover el punto de inserción al nuevo párrafo (final del documento)
        Set endRange = oDoc.Range(Start:=oDoc.Content.End - 1, End:=oDoc.Content.End - 1)
    End If
    
    endRange.Collapse wdCollapseStart ' Colapsar al inicio del rango (justo después de la tabla o en el nuevo párrafo)
    endRange.Select                  ' Seleccionar para mover el cursor del usuario

    ' --- Limpiar objetos VBA ---
    Set itemRange = Nothing
    Set oCell = Nothing
    Set oTable = Nothing
    Set originalRange = Nothing
    Set endRange = Nothing
    ' Set oDoc = Nothing ' No es estrictamente necesario para ActiveDocument

    ' Mensaje de finalización opcional (descomentar si se desea)
    ' MsgBox "Texto XML formateado como código en una tabla.", vbInformation, "Proceso Completado"
End Sub


Sub FormatearComoCodigoHTML()
    Dim selText As String
    Dim oTable As Table
    Dim oCell As Cell
    Dim oDoc As Document
    Dim rCelda As Range
    Dim etiquetasHTML As Variant
    Dim atributosHTML As Variant
    Dim simbolosHTML As Variant
    Dim itemHTML As Variant
    Dim originalRange As Range

    ' Obtener documento actual
    Set oDoc = ActiveDocument

    ' Verificar protección
    If oDoc.ReadOnlyRecommended Or oDoc.ProtectionType <> wdNoProtection Then
        MsgBox "El documento está protegido o es de solo lectura.", vbExclamation, "Error"
        Exit Sub
    End If

    ' Verificar texto seleccionado
    If Selection.Type = wdSelectionIP Or Len(Selection.Text) <= 1 Then
        MsgBox "Selecciona el texto HTML que deseas formatear.", vbInformation, "Selección vacía"
        Exit Sub
    End If

    ' Guardar texto seleccionado y rango original
    selText = Selection.Text
    Set originalRange = Selection.Range

    ' Listas de palabras clave y símbolos
    etiquetasHTML = Array("html", "head", "title", "meta", "link", "style", "script", "body", "div", "p", "span", "a", "img", "ul", "ol", "li", "table", "tr", "td", "th", "thead", "tbody", "tfoot", "form", "input", "textarea", "button")
    atributosHTML = Array("class", "id", "href", "src", "alt", "name", "type", "value", "placeholder", "method", "action", "style", "rel", "target", "width", "height")
    simbolosHTML = Array("<", ">", "</", "/>", "=", """", "'")

    ' Crear tabla y reemplazar selección
    Set oTable = oDoc.Tables.Add(Range:=originalRange, NumRows:=1, NumColumns:=1)
    Set oCell = oTable.Cell(1, 1)
    Set rCelda = oCell.Range
    rCelda.Text = selText
    rCelda.End = rCelda.End - 1 ' Eliminar marca de fin de celda

    ' Resaltar etiquetas
    For Each itemHTML In etiquetasHTML
        Call ResaltarTexto(rCelda, CStr(itemHTML), wdColorBlue, True)
    Next itemHTML

    ' Resaltar atributos
    For Each itemHTML In atributosHTML
        Call ResaltarTexto(rCelda, CStr(itemHTML), wdColorGray50, False)
    Next itemHTML

    ' Resaltar símbolos
    For Each itemHTML In simbolosHTML
        Call ResaltarTexto(rCelda, CStr(itemHTML), wdColorRed, False)
    Next itemHTML
End Sub

' ----------------------------------------------
' SUB auxiliar para resaltar palabras en un rango
' ----------------------------------------------
Sub ResaltarTexto(rangoBusqueda As Range, textoObjetivo As String, colorTexto As WdColor, subrayar As Boolean)
    Dim inicio As Long
    Dim encontrado As Range

    Set encontrado = rangoBusqueda.Duplicate
    With encontrado.Find
        .ClearFormatting
        .Text = textoObjetivo
        .Forward = True
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .Wrap = wdFindStop
        Do While .Execute
            encontrado.Font.color = colorTexto
            encontrado.Font.Bold = False
            encontrado.Font.Underline = IIf(subrayar, wdUnderlineSingle, wdUnderlineNone)
            encontrado.Collapse Direction:=wdCollapseEnd
        Loop
    End With
End Sub


' Sub auxiliar para resaltar palabras en un rango con opción de formato
Sub ResaltarTextoHTTP(rangoBusqueda As Range, textoObjetivo As String, colorTexto As Long, negrita As Boolean, subrayado As Boolean, Optional cursiva As Boolean = False)
    Dim encontrado As Range
    Set encontrado = rangoBusqueda.Duplicate

    With encontrado.Find
        .ClearFormatting
        .Text = textoObjetivo
        .Forward = True
        .Format = False
        .MatchCase = True
        .MatchWholeWord = True
        .Wrap = wdFindStop

        Do While .Execute
            encontrado.Font.color = colorTexto
            encontrado.Font.Bold = negrita
            encontrado.Font.Italic = cursiva
            If subrayado Then
                encontrado.Font.Underline = wdUnderlineSingle
            Else
                encontrado.Font.Underline = wdUnderlineNone
            End If
            encontrado.Collapse Direction:=wdCollapseEnd
        Loop
    End With
End Sub



' Sub principal para insertar y formatear solicitudes HTTP en una tabla
Sub FormatearSolicitudHTTP()
    Dim oDoc As Document
    Dim selText As String
    Dim originalRange As Range
    Dim oTable As Table
    Dim oCell As Cell
    Dim rCelda As Range
    Dim metodosHTTP As Variant
    Dim encabezadosHTTP As Variant
    Dim i As Long

    ' Obtener documento actual
    Set oDoc = ActiveDocument

    ' Verificar protección
    If oDoc.ReadOnlyRecommended Or oDoc.ProtectionType <> wdNoProtection Then
        MsgBox "El documento está protegido o es de solo lectura.", vbExclamation, "Error"
        Exit Sub
    End If

    ' Verificar texto seleccionado
    If Selection.Type = wdSelectionIP Or Len(Selection.Text) <= 1 Then
        MsgBox "Selecciona la solicitud HTTP que deseas formatear.", vbInformation, "Selección vacía"
        Exit Sub
    End If

    ' Guardar texto seleccionado y rango original
    selText = Selection.Text
    Set originalRange = Selection.Range

    ' Crear tabla y reemplazar selección
    Set oTable = oDoc.Tables.Add(Range:=originalRange, NumRows:=1, NumColumns:=1)
    Set oCell = oTable.Cell(1, 1)
    Set rCelda = oCell.Range
    rCelda.Text = selText
    rCelda.End = rCelda.End - 1 ' Eliminar marca de fin de celda
    
    ' Asegurar que el texto tenga color automático antes de aplicar formatos
rCelda.Font.color = wdColorAutomatic

    ' Definir métodos y encabezados HTTP
    metodosHTTP = Array("GET", "POST", "PUT", "DELETE", "OPTIONS", "HEAD", "PATCH")
    encabezadosHTTP = Array("Host", "User-Agent", "Accept", "Content-Type", "Content-Length", "Authorization", "Connection", "Accept-Encoding")

    ' Aplicar formato a métodos HTTP
    For i = LBound(metodosHTTP) To UBound(metodosHTTP)
        Call ResaltarTextoHTTP(rCelda, CStr(metodosHTTP(i)), wdColorBlue, True, True)
    Next i

    ' Aplicar formato a encabezados HTTP
    For i = LBound(encabezadosHTTP) To UBound(encabezadosHTTP)
        Call ResaltarTextoHTTP(rCelda, CStr(encabezadosHTTP(i)), wdColorDarkRed, True, False, True)
    Next i

    MsgBox "Solicitud HTTP formateada correctamente.", vbInformation
End Sub


' Sub auxiliar para resaltar URLs (http:// o https://)
Sub ResaltarURL(rangoBusqueda As Range)
    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    Dim startPos As Long, lengthURL As Long
    Dim rURL As Range
    
    ' Crear expresión regular para URL básica http/https
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "(http|https)://[^\s]+"
    regex.Global = True
    regex.IgnoreCase = True
    
    Set matches = regex.Execute(rangoBusqueda.Text)
    
    For Each match In matches
        startPos = match.FirstIndex + 1 ' +1 porque VBA es 1-based
        lengthURL = match.Length
        
        Set rURL = rangoBusqueda.Duplicate
        rURL.Start = rangoBusqueda.Start + startPos - 1
        rURL.End = rURL.Start + lengthURL
        
        With rURL.Font
            .color = wdColorDarkRed
            .Bold = True
            .Underline = wdUnderlineSingle
        End With
    Next
End Sub


Sub PalabrasClaveVerde_Corregida()

    Dim doc As Document
    Dim rng As Range          ' Rango de la selección original
    Dim searchRng As Range    ' Rango donde buscar (basado en la selección)
    Dim foundRng As Range     ' Rango específico de una coincidencia encontrada
    Dim regex As Object       ' Objeto para expresiones regulares
    Dim matches As Object     ' Colección de coincidencias
    Dim match As Object       ' Coincidencia individual
    Dim startPos As Long      ' Posición inicial de la coincidencia en el documento
    Dim endPos As Long        ' Posición final de la coincidencia en el documento
    Dim i As Long             ' Contador para el bucle de patrones
    Dim k As Long             ' Contador para el bucle de coincidencias (iteración inversa)

    ' --- CONFIGURACIÓN: PATRONES Y COLOR ---
    Const ignoreWord As String = "imgventk" ' Palabra a ignorar (no resaltar)

    ' --- Patrones (Expresiones Regulares) ---
    ' Usamos \b para asegurar que sean palabras completas (excepto en email e IP donde la estructura ya lo delimita)
    Const pattern1_IP As String = "\b(?:\d{1,3}\.){3}\d{1,3}\b"  ' IPs v4
    Const pattern2_Email As String = "[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}" ' Emails
    Const pattern3_Encoded As String = "\bEncoded\b"  ' Palabra "Encoded"
    Const pattern4_Payload As String = "\bPayload\b"  ' Palabra "Payload"
    Const pattern5_SpanishEncode As String = "\b(codificado|decodificamos)\b" ' Palabras "codificado" o "decodificamos"
    Const pattern6_SpecificVars As String = "\b(cveEncoded|cveDecoded)\b" ' Variables "cveEncoded" o "cveDecoded"

    ' --- Color de Resaltado ---
   Const colorHighlight As Integer = 4 ' Solo Verde para todos


    ' Verificar si hay texto seleccionado
    If Selection.Type = wdSelectionIP Or Selection.Type = wdSelectionNone Then
        MsgBox "Por favor, seleccione el texto donde desea resaltar las expresiones.", vbInformation
        Exit Sub
    End If

    Set doc = ActiveDocument
    Set rng = Selection.Range ' Rango original seleccionado
    ' Creamos un duplicado del rango de selección para realizar la búsqueda sin alterar el original
    Set searchRng = rng.Duplicate

    ' Crear y configurar el objeto de expresión regular
    On Error Resume Next ' Temporalmente ignorar errores si el objeto no se puede crear
    Set regex = CreateObject("VBScript.RegExp")
    If Err.Number <> 0 Then
        On Error GoTo 0 ' Restaurar manejo de errores normal
        MsgBox "Error al crear el objeto de Expresión Regular." & vbCrLf & _
               "Asegúrate de que 'VBScript Regular Expressions 5.5' esté disponible en el sistema.", vbCritical
        ' Liberar objetos antes de salir
        Set doc = Nothing: Set rng = Nothing: Set searchRng = Nothing
        Exit Sub
    End If
    On Error GoTo 0 ' Restaurar manejo de errores normal

    With regex
        .Global = True      ' Buscar todas las ocurrencias, no solo la primera
        .IgnoreCase = True  ' Ignorar mayúsculas/minúsculas
        .MultiLine = True   ' Tratar ^ y $ como inicio/fin de línea (aunque no es crucial para estos patrones)
    End With

    ' --- Agrupar los patrones en un array para iterar ---
    ' Nota: Solo definimos los patrones que realmente estamos usando
    Dim patterns(1 To 6) As String
    patterns(1) = pattern1_IP
    patterns(2) = pattern2_Email
    patterns(3) = pattern3_Encoded
    patterns(4) = pattern4_Payload
    patterns(5) = pattern5_SpanishEncode
    patterns(6) = pattern6_SpecificVars

    ' --- Procesar cada patrón ---
    For i = LBound(patterns) To UBound(patterns)
        regex.Pattern = patterns(i) ' Establecer el patrón actual

        ' Ejecutar la búsqueda en el texto del rango seleccionado
        If regex.Test(searchRng.Text) Then ' Optimización: Ejecutar solo si hay al menos una coincidencia
            Set matches = regex.Execute(searchRng.Text)

            ' Iterar sobre las coincidencias EN ORDEN INVERSO
            ' Esto es importante porque al aplicar formato (resaltado), las posiciones
            ' de las coincidencias posteriores podrían cambiar si iteramos hacia adelante.
            For k = matches.Count - 1 To 0 Step -1
                Set match = matches(k)

                ' Comprobar si la coincidencia es la palabra a ignorar (comparando en minúsculas)
                If LCase(match.Value) <> LCase(ignoreWord) Then
                    ' Calcular las posiciones de inicio y fin de la coincidencia
                    ' DENTRO DEL DOCUMENTO COMPLETO, usando el inicio del rango de búsqueda como base.
                    startPos = searchRng.Start + match.FirstIndex
                    endPos = startPos + match.Length

                    ' Crear un rango específico para esta coincidencia dentro del documento
                    Set foundRng = doc.Range(startPos, endPos)

                    ' Aplicar el color de resaltado
                    ' Solo aplicar si aún no tiene ESE MISMO color (evita trabajo innecesario, aunque no es estrictamente necesario)
                    If foundRng.HighlightColorIndex <> colorHighlight Then
                         foundRng.HighlightColorIndex = colorHighlight
                    End If

                    ' Liberar el rango de la coincidencia encontrada para la siguiente iteración
                    Set foundRng = Nothing
                End If
            Next k ' Siguiente coincidencia (en orden inverso)

            ' Liberar la colección de coincidencias para el patrón actual
            Set matches = Nothing
        End If ' Fin de If regex.Test
    Next i ' Siguiente patrón

    ' --- Limpieza final de objetos ---
    Set regex = Nothing
    Set match = Nothing
    Set rng = Nothing
    Set searchRng = Nothing
    Set doc = Nothing

    ' Opcional: Mensaje de finalización
    ' MsgBox "Resaltado completado.", vbInformation

End Sub

