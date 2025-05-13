Attribute VB_Name = "NewMacros"
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



Sub AjustarFormatoColumnasTablaVulnes()
    Dim tbl As table
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
        For r = 2 To tbl.Rows.count
            tbl.Rows(r).Range.Font.Size = 10
        Next r

        ' Ajustar anchos de columnas (fijos)
        Dim i As Integer
        For i = 1 To 4
            If i <= tbl.Columns.count Then
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
    selRange.Font.Color = colorPorDefecto

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
                word.Font.Color = colorHTML
                colored = True
                Exit For ' Salir del bucle de HTML si se encuentra coincidencia
            End If
        Next keyword
        If colored Then GoTo NextWord ' Ir a la siguiente palabra si ya se coloreó

        ' Comprobar CSS (solo si no es HTML)
        For Each keyword In cssKeywords
            If wordText = keyword Then
                word.Font.Color = colorCSS
                colored = True
                Exit For
            End If
        Next keyword
        If colored Then GoTo NextWord

        ' Comprobar JS (solo si no es HTML ni CSS)
        For Each keyword In jsKeywords
            If wordText = keyword Then
                word.Font.Color = colorJS
                colored = True
                Exit For
            End If
        Next keyword
        If colored Then GoTo NextWord

        ' Comprobar Valores comunes (solo si no es ninguna de las anteriores)
        For Each keyword In valueKeywords
            If wordText = keyword Then
                word.Font.Color = colorValores
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
            For k = matches.count - 1 To 0 Step -1
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

