Attribute VB_Name = "NewMacros"
Sub NegritaPalabrasClave_Robusta_MultiArray_Corregido()

    ' Declaraci�n de variables
    Dim palabrasClaveParte1 As Variant
    Dim palabrasClaveParte2 As Variant
    Dim palabrasClaveParte3 As Variant
    ' Si necesitas m�s partes, decl�ralas aqu�:
    ' Dim palabrasClaveParte4 As Variant
    ' Dim palabrasClaveParte5 As Variant
    ' ... y as� sucesivamente

    Dim todasLasPalabrasClaveList As Object ' Usaremos un ArrayList para combinar
    Dim palabrasClaveOrdenadas As Variant
    Dim palabra As Variant ' Para iterar sobre las palabras clave
    Dim item As Variant    ' Para iterar sobre el ArrayList
    Dim rangoAProcesar As Range
    Dim i As Long          ' Contador para el bucle de ordenaci�n
    Dim temp As String     ' Variable temporal para el intercambio en la ordenaci�n
    Dim swapped As Boolean ' Bandera para el bucle de ordenaci�n

    ' --- INICIO: Definici�n de palabras clave en fragmentos de Arrays ---
    ' Distribuye tus palabras clave aqu�.
    ' Aseg�rate de que cada array individual no exceda el l�mite de continuaci�n de l�nea.
    ' Ejemplo con una porci�n de tu lista:

palabrasClaveParte1 = Array( _
    "XSS", "Stored XSS", "DOM-based XSS", _
    "Session Hijacking", "Phishing", "CSP", _
    "Validaci�n de entradas", "Sanitizaci�n", "Interceptaci�n", _
    "TLS 1.0", "Protocolo d�bil", "Man-in-the-Middle" _
)

palabrasClaveParte2 = Array( _
    "Malware", "Explotaci�n", "TLS 1.1", "controles de restricci�n adecuados", _
    "Downgrade Attack", "Tr�fico TLS", "Sweet32", _
    "Wireshark", "Tshark", "tcpdump", _
    "Gesti�n de vulnerabilidades", "Seguridad de aplicaciones web", "Divulgaci�n de informaci�n" _
)

palabrasClaveParte3 = Array( _
    "Fuga de informaci�n", "Hardening", "HTTP Headers", _
 "Autodiscover", "X-Frame-Options", "HttpOnly", _
  "SSL Stripping", "Componentes vulnerables", _
    "OWASP DependencyCheck", "Subresource Integrity", "Clickjacking", _
    "Cookies HttpOnly", "Cookies Secure", "Fingerprinting", _
    "CVE", _
    "Metodolog�a de pentesting", "Seguridad en la nube", "Acceso no autorizado", _
    "Credenciales", "Confidencialidad", "Remediaci�n" _
)


    Set todasLasPalabrasClaveList = CreateObject("System.Collections.ArrayList")

    ' A�adir elementos de cada parte del array a la lista combinada
    For Each palabra In palabrasClaveParte1
        If Len(Trim(CStr(palabra))) > 0 Then ' Asegurar que no se a�adan strings vac�os
            todasLasPalabrasClaveList.Add Trim(CStr(palabra))
        End If
    Next palabra

    For Each palabra In palabrasClaveParte2
        If Len(Trim(CStr(palabra))) > 0 Then
            todasLasPalabrasClaveList.Add Trim(CStr(palabra))
        End If
    Next palabra

    For Each palabra In palabrasClaveParte3
        If Len(Trim(CStr(palabra))) > 0 Then
            todasLasPalabrasClaveList.Add Trim(CStr(palabra))
        End If
    Next palabra

    ' Si a�adiste m�s partes (palabrasClaveParte4, etc.), a�ade bucles para ellas aqu�:
    ' For Each palabra In palabrasClaveParte4
    '    If Len(Trim(CStr(palabra))) > 0 Then
    '        todasLasPalabrasClaveList.Add Trim(CStr(palabra))
    '    End If
    ' Next palabra


    ' --- Ordenar la lista combinada por longitud de cadena, de mayor a menor ---
    ' Esto es importante si MatchWholeWord = False
    If todasLasPalabrasClaveList.count > 1 Then
        Do
            swapped = False
            For i = 0 To todasLasPalabrasClaveList.count - 2
                If Len(todasLasPalabrasClaveList(i)) < Len(todasLasPalabrasClaveList(i + 1)) Then
                    temp = todasLasPalabrasClaveList(i)
                    todasLasPalabrasClaveList(i) = todasLasPalabrasClaveList(i + 1)
                    todasLasPalabrasClaveList(i + 1) = temp
                    swapped = True
                End If
            Next i
        Loop While swapped
    End If
    palabrasClaveOrdenadas = todasLasPalabrasClaveList.ToArray() ' Convertir ArrayList a Array VBA

    ' --- Determinar el rango a procesar ---
    If Selection.Type = wdSelectionIP Or Selection.Type = wdNoSelection Then
        MsgBox "Por favor, seleccione el texto donde desea aplicar la negrita.", vbExclamation
        Exit Sub ' Salir si no hay selecci�n v�lida
    End If
    Set rangoAProcesar = Selection.Range
    ' Para procesar todo el documento en lugar de la selecci�n, descomenta la siguiente l�nea
    ' y comenta la l�nea anterior:
    ' Set rangoAProcesar = ActiveDocument.Content
    
    ' Desactivar actualizaci�n de pantalla para mejorar rendimiento
    Application.ScreenUpdating = False
    
    ' Preparar el objeto Find para el rango a procesar
    Dim findObj As Find
    Set findObj = rangoAProcesar.Find

    ' --- Bucle principal para buscar y reemplazar cada palabra clave ---
    For Each palabra In palabrasClaveOrdenadas
        ' 'palabra' ya est� trimeada y no est� vac�a gracias al preprocesamiento
        
        With findObj
            .ClearFormatting
            .Text = palabra ' La palabra clave actual a buscar
            .Replacement.ClearFormatting
            .Replacement.Font.Bold = True ' Aplicar negrita al reemplazo
            .Forward = True               ' Buscar hacia adelante
            .Wrap = wdFindContinue        ' Continuar buscando en todo el rango especificado
            .Format = True                ' Considerar el formato en la b�squeda (aunque aqu� se limpia)
            .MatchCase = False            ' Ignorar may�sculas/min�sculas
            .MatchWholeWord = False       ' No buscar solo palabras completas (la ordenaci�n ayuda aqu�)
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            
            .Execute Replace:=wdReplaceAll ' Reemplazar todas las ocurrencias
        End With
    Next palabra
    
    ' --- Restaurar la configuraci�n del objeto Find de la Selecci�n a un estado m�s neutral ---
    ' Esto es para evitar que la macro afecte las b�squedas manuales posteriores del usuario.
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "" ' Limpiar el texto de b�squeda
        .Forward = True
        .Wrap = wdFindAsk ' O wdFindStop, seg�n preferencia para b�squedas manuales
        .Format = False   ' Generalmente False para b�squedas manuales sin formato espec�fico
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        ' Asegurarse de que no haya formato de reemplazo persistente
        .Replacement.Font.Bold = False
        ' Si se modificaron otras propiedades de .Replacement.Font, restaurarlas tambi�n.
    End With

    ' Reactivar actualizaci�n de pantalla
    Application.ScreenUpdating = True
    
    ' Mensaje de finalizaci�n
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

        ' Desactivar ajuste autom�tico de columnas
        tbl.AllowAutoFit = False

        ' Centrar la tabla en la p�gina
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











Sub NegritaPalabrasClave()
    Dim palabrasTexto As String
    Dim palabrasClave As Variant
    Dim palabra As Variant
    Dim rango As Range
    

palabrasTexto = "Cross-Site Scripting (XSS), XSS Reflected, Reflected XSS, XSS Persistent, Stored XSS, DOM-based XSS, Inyecci�n, Scripts maliciosos, Malicious Scripts, Ejecuci�n de acciones en nombre del usuario, Robo de cookies de sesi�n, Session Hijacking, Redirecci�n a sitios fraudulentos, Phishing, Content Security Policy (CSP), Validaci�n, Sanitizaci�n, Manipulaci�n de sesiones de usuario, Formulario malicioso, TLS 1.0, Protocolo d�bil, Algoritmos de cifrado, Interceptaci�n, Hombre en el medio, Escucha pasiva, Malware, Explotaci�n,TLS 1.1, downgrade, Interceptar tr�fico cifrado, Atacante, Wireshark, Tshark, tcpdump, Tr�fico TLS, Sweet32, Colisiones, CBC, (Cipher Block Chaining)"


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
    Const colorCSS As Long = &H228B22      ' Verde m�s claro para CSS (ej. propiedades)
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
    valueKeywords = Array("true", "false", "null", "undefined", "auto", "inherit", "initial", "unset", "none", "block", "inline", "inline-block", "flex", "grid", "absolute", "relative", "fixed", "static", "solid", "dotted", "dashed", "double", "hidden", "visible", "bold", "italic", "normal") ' A�adidos valores comunes

    ' ----- L�gica Principal -----
    ' 1. Verificar si hay texto seleccionado
    If Selection.Type = wdSelectionIP Or Selection.Type = wdSelectionNone Then
        MsgBox "Por favor, primero seleccione el bloque de c�digo HTML/CSS/JS que desea formatear.", vbInformation, "Selecci�n Requerida"
        Exit Sub
    End If

    Set selRange = Selection.Range

    ' 2. Restablecer el color de toda la selecci�n a negro (o el color por defecto)
    ' Esto asegura que solo las palabras clave identificadas cambien de color.
    selRange.Font.Color = colorPorDefecto

    ' 3. Iterar sobre cada "palabra" en la selecci�n
    ' Usamos Trim() para quitar espacios al inicio/final que a veces incluye .Words
    ' Usamos LCase() para hacer la comparaci�n insensible a may�sculas/min�sculas
    For Each word In selRange.Words
        Dim wordText As String
        wordText = LCase(Trim(word.Text))
        
        ' Optimizaci�n: Si la palabra est� vac�a despu�s de Trim, saltarla.
        If Len(wordText) = 0 Then GoTo NextWord

        ' Bandera para saber si ya se colore� la palabra
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
        If colored Then GoTo NextWord ' Ir a la siguiente palabra si ya se colore�

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
        ' No necesita GoTo aqu� ya que es la �ltima comprobaci�n de palabras clave

' Etiqueta para saltar a la siguiente iteraci�n del bucle principal
NextWord:
    Next word

    ' ----- Limpieza (Buenas pr�cticas) -----
    Set word = Nothing
    Set selRange = Nothing

    ' Mensaje Opcional de finalizaci�n
    ' MsgBox "Formato de sintaxis aplicado a la selecci�n.", vbInformation

End Sub



Sub PalabrasClaveVerde_Corregida()

    Dim doc As Document
    Dim rng As Range          ' Rango de la selecci�n original
    Dim searchRng As Range    ' Rango donde buscar (basado en la selecci�n)
    Dim foundRng As Range     ' Rango espec�fico de una coincidencia encontrada
    Dim regex As Object       ' Objeto para expresiones regulares
    Dim matches As Object     ' Colecci�n de coincidencias
    Dim match As Object       ' Coincidencia individual
    Dim startPos As Long      ' Posici�n inicial de la coincidencia en el documento
    Dim endPos As Long        ' Posici�n final de la coincidencia en el documento
    Dim i As Long             ' Contador para el bucle de patrones
    Dim k As Long             ' Contador para el bucle de coincidencias (iteraci�n inversa)

    ' --- CONFIGURACI�N: PATRONES Y COLOR ---
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
    ' Creamos un duplicado del rango de selecci�n para realizar la b�squeda sin alterar el original
    Set searchRng = rng.Duplicate

    ' Crear y configurar el objeto de expresi�n regular
    On Error Resume Next ' Temporalmente ignorar errores si el objeto no se puede crear
    Set regex = CreateObject("VBScript.RegExp")
    If Err.Number <> 0 Then
        On Error GoTo 0 ' Restaurar manejo de errores normal
        MsgBox "Error al crear el objeto de Expresi�n Regular." & vbCrLf & _
               "Aseg�rate de que 'VBScript Regular Expressions 5.5' est� disponible en el sistema.", vbCritical
        ' Liberar objetos antes de salir
        Set doc = Nothing: Set rng = Nothing: Set searchRng = Nothing
        Exit Sub
    End If
    On Error GoTo 0 ' Restaurar manejo de errores normal

    With regex
        .Global = True      ' Buscar todas las ocurrencias, no solo la primera
        .IgnoreCase = True  ' Ignorar may�sculas/min�sculas
        .MultiLine = True   ' Tratar ^ y $ como inicio/fin de l�nea (aunque no es crucial para estos patrones)
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

    ' --- Procesar cada patr�n ---
    For i = LBound(patterns) To UBound(patterns)
        regex.Pattern = patterns(i) ' Establecer el patr�n actual

        ' Ejecutar la b�squeda en el texto del rango seleccionado
        If regex.Test(searchRng.Text) Then ' Optimizaci�n: Ejecutar solo si hay al menos una coincidencia
            Set matches = regex.Execute(searchRng.Text)

            ' Iterar sobre las coincidencias EN ORDEN INVERSO
            ' Esto es importante porque al aplicar formato (resaltado), las posiciones
            ' de las coincidencias posteriores podr�an cambiar si iteramos hacia adelante.
            For k = matches.count - 1 To 0 Step -1
                Set match = matches(k)

                ' Comprobar si la coincidencia es la palabra a ignorar (comparando en min�sculas)
                If LCase(match.Value) <> LCase(ignoreWord) Then
                    ' Calcular las posiciones de inicio y fin de la coincidencia
                    ' DENTRO DEL DOCUMENTO COMPLETO, usando el inicio del rango de b�squeda como base.
                    startPos = searchRng.Start + match.FirstIndex
                    endPos = startPos + match.Length

                    ' Crear un rango espec�fico para esta coincidencia dentro del documento
                    Set foundRng = doc.Range(startPos, endPos)

                    ' Aplicar el color de resaltado
                    ' Solo aplicar si a�n no tiene ESE MISMO color (evita trabajo innecesario, aunque no es estrictamente necesario)
                    If foundRng.HighlightColorIndex <> colorHighlight Then
                         foundRng.HighlightColorIndex = colorHighlight
                    End If

                    ' Liberar el rango de la coincidencia encontrada para la siguiente iteraci�n
                    Set foundRng = Nothing
                End If
            Next k ' Siguiente coincidencia (en orden inverso)

            ' Liberar la colecci�n de coincidencias para el patr�n actual
            Set matches = Nothing
        End If ' Fin de If regex.Test
    Next i ' Siguiente patr�n

    ' --- Limpieza final de objetos ---
    Set regex = Nothing
    Set match = Nothing
    Set rng = Nothing
    Set searchRng = Nothing
    Set doc = Nothing

    ' Opcional: Mensaje de finalizaci�n
    ' MsgBox "Resaltado completado.", vbInformation

End Sub

