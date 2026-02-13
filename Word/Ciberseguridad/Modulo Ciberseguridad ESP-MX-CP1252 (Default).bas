Attribute VB_Name = "Ciber"
Option Explicit


' ==============================================================================
' PARTE 1: AUTOMATIZACIÓN (Para que funcione en Módulos)
' ==============================================================================

Sub CYB_001_AutoOpen()
    Call CYB_003_AddSeverityContextMenu
End Sub

Sub CYB_002_AutoClose()
    Call CYB_004_RemoveSeverityContextMenu
    If Not ActiveDocument.Saved Then ActiveDocument.Save
End Sub

' ==============================================================================
' PARTE 2: CREAR EL MENÚ
' ==============================================================================

Sub CYB_003_AddSeverityContextMenu()
    Dim menuBar As CommandBar
    Dim newMenu As CommandBarPopup
    Dim menuItem As CommandBarButton
    
    Set menuBar = Application.CommandBars("Table Text")
    
    On Error Resume Next
    menuBar.Controls("Establecer Severidad").Delete
    On Error GoTo 0
    
    Set newMenu = menuBar.Controls.Add(Type:=msoControlPopup, Before:=1)
    newMenu.Caption = "Establecer Severidad"
    newMenu.Tag = "MenuSeveridad"
    
    ' --- BOTONES ---
    
    ' CRÍTICA
    Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
    menuItem.Caption = "1 - CRÍTICA"
    menuItem.FaceId = 66
    menuItem.OnAction = "CYB_005_ColorearCritica"
    
    ' ALTA
    Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
    menuItem.Caption = "2 - ALTA"
    menuItem.FaceId = 207
    menuItem.OnAction = "CYB_006_ColorearAlta"
    
    ' MEDIA
    Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
    menuItem.Caption = "3 - MEDIA"
    menuItem.FaceId = 139
    menuItem.OnAction = "CYB_007_ColorearMedia"
    
    ' BAJA
    Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
    menuItem.Caption = "4 - BAJA"
    menuItem.FaceId = 138
    menuItem.OnAction = "CYB_008_ColorearBaja"

    ' QUITAR COLOR
    Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
    menuItem.Caption = "Quitar Color"
    menuItem.FaceId = 1845
    menuItem.OnAction = "CYB_009_QuitarColor"

End Sub

Sub CYB_004_RemoveSeverityContextMenu()
    On Error Resume Next
    Application.CommandBars("Table Text").Controls("Establecer Severidad").Delete
End Sub

' ==============================================================================
' PARTE 3: MACROS QUE CAMBIAN TEXTO Y COLOR
' ==============================================================================

Sub CYB_005_ColorearCritica()
    If Selection.Information(wdWithInTable) Then
        With Selection.Cells(1)
            ' 1. Escribe el texto en MAYÚSCULAS
            ' (Se usa un caracter especial para no borrar la celda entera)
            .Range.Text = "CRÍTICA"
            
            ' 2. Aplica colores (Fondo Rojo / Letra Blanca)
            .Shading.BackgroundPatternColor = wdColorRed
            .Range.Font.color = wdColorWhite
            .Range.Font.Bold = True ' Opcional: Pone negrita para resaltar más
        End With
    End If
End Sub

Sub CYB_006_ColorearAlta()
    If Selection.Information(wdWithInTable) Then
        With Selection.Cells(1)
            ' 1. Escribe texto
            .Range.Text = "ALTA"
            
            ' 2. Aplica colores (Fondo Naranja / Letra Blanca)
            .Shading.BackgroundPatternColor = RGB(255, 69, 0)
            .Range.Font.color = wdColorWhite
            .Range.Font.Bold = True
        End With
    End If
End Sub

Sub CYB_007_ColorearMedia()
    If Selection.Information(wdWithInTable) Then
        With Selection.Cells(1)
            ' 1. Escribe texto
            .Range.Text = "MEDIA"
            
            ' 2. Aplica colores (Fondo Amarillo / Letra NEGRA)
            .Shading.BackgroundPatternColor = wdColorYellow
            .Range.Font.color = wdColorBlack
            .Range.Font.Bold = True
        End With
    End If
End Sub

Sub CYB_008_ColorearBaja()
    If Selection.Information(wdWithInTable) Then
        With Selection.Cells(1)
            ' 1. Escribe texto
            .Range.Text = "BAJA"
            
            ' 2. Aplica colores (Fondo Verde / Letra NEGRA)
            .Shading.BackgroundPatternColor = wdColorBrightGreen
            .Range.Font.color = wdColorBlack
            .Range.Font.Bold = True
        End With
    End If
End Sub

Sub CYB_009_QuitarColor()
    ' NOTA: Este solo quita el color, NO borra el texto,
    ' por si te equivocaste y quieres mantener lo escrito.
    On Error Resume Next
    If Selection.Information(wdWithInTable) = True Then
        With Selection.Cells(1)
            .Shading.BackgroundPatternColor = wdColorWhite
            .Range.Font.color = wdColorAutomatic
            .Range.Font.Bold = False
        End With
    End If
End Sub


Sub CYB_010_FormatearTablaVulnerabilidadesAvanzado()

    Dim tbl As table
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
    If tbl.Rows.count >= 1 Then
        For j = 1 To tbl.Rows(1).Cells.count
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
    If tbl.Rows.count >= 1 Then
        With tbl.Rows(1)
            .Shading.BackgroundPatternColor = HEADER_BLUE
            .Range.Font.ColorIndex = wdWhite
            .Range.Font.Bold = True
            .Cells.VerticalAlignment = wdCellAlignVerticalCenter ' Centrado Vertical
            Dim cl As cell
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
    If tbl.Rows.count >= 2 Then
        For i = 2 To tbl.Rows.count
            
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
                If tbl.Rows(i).Cells.count >= severityColIndex Then
                    With tbl.cell(i, severityColIndex)
                        If .Range.Characters.count > 0 Then
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
                If tbl.Rows(i).Cells.count >= estadoColIndex Then
                    With tbl.cell(i, estadoColIndex)
                        If .Range.Characters.count > 0 Then
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
    
' --- 5. Aplicar Bordes ---

Dim r As row, c As cell

' --- Quitar bordes antiguos ---
With tbl
    ' Limpia cualquier borde previo en tabla
    With .Borders
        .InsideLineStyle = wdLineStyleNone
        .OutsideLineStyle = wdLineStyleNone
    End With

    For Each r In .Rows
        For Each c In r.Cells
            c.Borders(wdBorderTop).LineStyle = wdLineStyleNone
            c.Borders(wdBorderBottom).LineStyle = wdLineStyleNone
            c.Borders(wdBorderLeft).LineStyle = wdLineStyleNone
            c.Borders(wdBorderRight).LineStyle = wdLineStyleNone
        Next c
    Next r
End With



' --- Aplicar bordes visibles a todas las celdas ---
For Each r In tbl.Rows
    For Each c In r.Cells
        With c.Borders
            .Item(wdBorderTop).LineStyle = wdLineStyleSingle
            .Item(wdBorderBottom).LineStyle = wdLineStyleSingle
            .Item(wdBorderLeft).LineStyle = wdLineStyleSingle
            .Item(wdBorderRight).LineStyle = wdLineStyleSingle
            
            ' Color de los bordes
            .Item(wdBorderTop).color = wdColorAutomatic
            .Item(wdBorderBottom).color = wdColorAutomatic
            .Item(wdBorderLeft).color = wdColorAutomatic
            .Item(wdBorderRight).color = wdColorAutomatic
        End With
    Next c
Next r


' --- Aplicar nuevo formato uniforme ---
With tbl
    .TopPadding = CentimetersToPoints(0.03)
    .BottomPadding = CentimetersToPoints(0.03)
    .LeftPadding = CentimetersToPoints(0.03)
    .RightPadding = CentimetersToPoints(0.03)
    .Spacing = 0
    .AllowAutoFit = True
    .Rows.Alignment = wdAlignRowCenter
End With



tbl.Rows.Alignment = wdAlignRowCenter


' Abrir el diálogo de opciones de tabla y apagar "Allow spacing between cells"
' Selecciona toda la tabla para que el diálogo se aplique a ella
tbl.Select

' Abrir el diálogo de opciones de tabla y apagar "Allow spacing between cells"
With Dialogs(wdDialogTableTableOptions)
    .AllowSpacing = False
    .Execute
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

Sub CYB_011_NegritaPalabrasClave_Robusta_MultiArray_Corregido_CVECompleto()

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

    ' Ordenar por longitud descendente para evitar conflictos
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

    palabrasClaveOrdenadas = todasLasPalabrasClaveList.ToArray()

    ' Definir el rango a procesar
    If Selection.Type = wdSelectionIP Or Selection.Type = wdNoSelection Then
        MsgBox "Por favor, seleccione el texto donde desea aplicar la negrita.", vbExclamation
        Exit Sub
    End If
    Set rangoAProcesar = Selection.Range

    Application.ScreenUpdating = False

    ' Primero negrita para palabras clave
    For Each palabra In palabrasClaveOrdenadas
        Set findObj = rangoAProcesar.Duplicate.Find
        With findObj
            .ClearFormatting
            .Text = palabra
            .Replacement.ClearFormatting
            .Replacement.Text = "^&"
            .Replacement.Font.Bold = True
            .Forward = True
            .Wrap = wdFindStop
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .Execute Replace:=wdReplaceAll
        End With
    Next palabra

    ' Luego negrita para **CVE completas** usando comodines
    Set findObj = rangoAProcesar.Duplicate.Find
    With findObj
        .ClearFormatting
        .Text = "CVE-[0-9]{4}-[0-9]{4,7}" ' patrón de CVE
        .Replacement.ClearFormatting
        .Replacement.Text = "^&"
        .Replacement.Font.Bold = True
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    Application.ScreenUpdating = True

    MsgBox "Palabras clave y CVE completas en negrita aplicadas.", vbInformation

End Sub




Sub CYB_012_AjustarFormatoColumnasTablaVulnes()
    ' --- DECLARACIÓN DE VARIABLES PRINCIPALES ---
    Dim tbl As word.table
    Dim col As word.Column
    Dim headerText As String
    Dim severidadColIndex As Long
    Dim r As Long
    Dim c As Long
    
    ' --- CONSTANTES ---
    Const COLOR_FILA_PAR As Long = wdColorGray10

    ' --- INICIO DE LA MACRO ---
    ' 1. VERIFICAR QUE EL CURSOR ESTÉ EN UNA TABLA
    If Selection.Information(wdWithInTable) Then
        Set tbl = Selection.Tables(1)

        ' --- 2. CONFIGURACIÓN GENERAL DE LA TABLA ---
        tbl.AllowAutoFit = False
        tbl.AutoFitBehavior wdAutoFitFixed
        tbl.Rows.AllowBreakAcrossPages = True
        tbl.Rows.Alignment = wdAlignRowCenter
        tbl.Range.Cells.VerticalAlignment = wdCellAlignVerticalCenter

        ' --- 3. FORMATO DE FUENTE INICIAL ---
        With tbl.Rows(1).Range.Font
            .Size = 11
            .Bold = True
        End With
        
        If tbl.Rows.count > 1 Then
            For r = 2 To tbl.Rows.count
                With tbl.Rows(r).Range.Font
                    .Size = 10
                    .Bold = False
                End With
            Next r
        End If

        ' --- 4. AJUSTE DINÁMICO DE ANCHO DE COLUMNAS ---
        severidadColIndex = 0
        For Each col In tbl.Columns
            headerText = Trim(Replace(col.Cells(1).Range.Text, Chr(13) & Chr(7), ""))
            If InStr(1, LCase(headerText), "severidad") > 0 Then
                severidadColIndex = col.Index
            End If
            
            Select Case True
                Case InStr(1, LCase(headerText), "severidad") > 0
                    col.PreferredWidth = CentimetersToPoints(2.8)
                Case InStr(1, LCase(headerText), "vulnerabilidad") > 0
                    col.PreferredWidth = CentimetersToPoints(6)
                Case InStr(1, LCase(headerText), "afectado") > 0
                    col.PreferredWidth = CentimetersToPoints(7)
                Case Else
                    col.PreferredWidth = CentimetersToPoints(5)
            End Select
            col.PreferredWidthType = wdPreferredWidthPoints
        Next col

        ' --- 4.1. AJUSTAR ANCHO TOTAL DE LA TABLA ---
        Dim anchoTotal As Single
        anchoTotal = 0
        For Each col In tbl.Columns
            anchoTotal = anchoTotal + col.PreferredWidth
        Next col
        tbl.PreferredWidthType = wdPreferredWidthPoints
        tbl.PreferredWidth = anchoTotal
        
        ' --- 5. APLICACIÓN DE COLORES Y FORMATO CELDA POR CELDA ---
        For r = 2 To tbl.Rows.count
            For c = 1 To tbl.Columns.count
                Dim currentCell As cell
                Set currentCell = tbl.cell(r, c)
                
                ' --- INICIO DEL IF/ELSE PRINCIPAL ---
                If c = severidadColIndex And severidadColIndex > 0 Then
                    ' --- BLOQUE IF: ESTAMOS EN LA COLUMNA DE SEVERIDAD ---
                    Dim cellText As String, bgColor As Long, fontColor As WdColor
                    Dim score As Double, colorAplicar As Boolean
                    cellText = Trim(Replace(currentCell.Range.Text, Chr(13) & Chr(7), ""))
                    colorAplicar = False
                    
                    If IsNumeric(cellText) Then
                        score = CDbl(cellText)
                        ' --- SINTAXIS DE BLOQUE CORREGIDA ---
                        If score >= 9# Then
                            bgColor = RGB(&H70, &H30, &HA0)
                            fontColor = wdColorWhite
                            colorAplicar = True
                        ElseIf score >= 7# Then
                            bgColor = RGB(255, 0, 0)
                            fontColor = wdColorWhite
                            colorAplicar = True
                        ElseIf score >= 4# Then
                            bgColor = RGB(255, 255, 0)
                            fontColor = wdColorBlack
                            colorAplicar = True
                        ElseIf score >= 0.1 Then
                            bgColor = RGB(0, 176, 80)
                            fontColor = wdColorWhite
                            colorAplicar = True
                        End If
                    Else
                        ' --- SINTAXIS DE BLOQUE CORREGIDA ---
                        Select Case UCase(cellText)
                            Case "CRÍTICA", "CRITICAL"
                                bgColor = RGB(&H70, &H30, &HA0)
                                fontColor = wdColorWhite
                                colorAplicar = True
                            Case "ALTA", "HIGH"
                                bgColor = RGB(255, 0, 0)
                                fontColor = wdColorWhite
                                colorAplicar = True
                            Case "MEDIA", "MEDIUM"
                                bgColor = RGB(255, 255, 0)
                                fontColor = wdColorBlack
                                colorAplicar = True
                            Case "BAJA", "LOW"
                                bgColor = RGB(0, 176, 80)
                                fontColor = wdColorWhite
                                colorAplicar = True
                        End Select
                    End If
                    
                    If colorAplicar Then
                        currentCell.Shading.BackgroundPatternColor = bgColor
                        With currentCell.Range.Font
                            .color = fontColor
                            .Bold = True
                        End With
                        currentCell.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                    Else
                        currentCell.Shading.BackgroundPatternColor = wdColorWhite
                        With currentCell.Range.Font
                            .color = wdColorAutomatic
                            .Bold = False
                        End With
                        currentCell.Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
                    End If

                Else
                    ' --- BLOQUE ELSE: ESTAMOS EN CUALQUIER OTRA COLUMNA ---
                    If r Mod 2 = 0 Then
                        currentCell.Shading.BackgroundPatternColor = COLOR_FILA_PAR
                    Else
                        currentCell.Shading.BackgroundPatternColor = wdColorWhite
                    End If
                    
                    With currentCell.Range.Font
                        .color = wdColorAutomatic
                        .Bold = False
                    End With

                    currentCell.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                    
                End If ' --- FIN DEL IF/ELSE PRINCIPAL ---

            Next c
        Next r

    Else
        ' Mensaje de error si no se selecciona una tabla
        MsgBox "Por favor, coloca el cursor dentro de la tabla que deseas formatear.", vbExclamation, "Tabla no seleccionada"
    End If
End Sub



Sub CYB_013_NegritaPalabrasClave()
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



Sub CYB_014_InsertarBloqueCodigoFormateado()
    If Selection.Type = wdSelectionIP Then
        MsgBox "Por favor, seleccione el código que desea convertir en bloque de código.", vbInformation
        Exit Sub
    End If

    Dim textoCodigo As String
    textoCodigo = Selection.Range.Text

    ' Insertar tabla contenedora
    Dim tabla As table
    Dim rngInsertar As Range
    Set rngInsertar = Selection.Range
    rngInsertar.Collapse wdCollapseEnd

    Set tabla = rngInsertar.Tables.Add(Range:=rngInsertar, NumRows:=1, NumColumns:=1)
    With tabla
        .Borders.Enable = False
        .Shading.BackgroundPatternColor = RGB(240, 240, 240) ' Gris claro
        .Rows.LeftIndent = 0
        .PreferredWidthType = wdPreferredWidthPercent
        .PreferredWidth = 100
        .AllowAutoFit = False
    End With

    ' Insertar el texto en la celda
    Dim celda As Range
    Set celda = tabla.cell(1, 1).Range
    celda.Text = textoCodigo
    celda.End = celda.End - 1 ' Quitar el carácter de fin de celda

    ' Aplicar formato base
    With celda.Font
        .Name = "Consolas"
        .Size = 10
        .color = RGB(0, 0, 0)
        .Bold = False
        .Italic = False
    End With

    ' Formatear sintaxis en el contenido de la celda
    Call CYB_016_FormatearCodigoEnRango(celda)

    Selection.Delete
End Sub

Option Explicit

Sub CYB_015_ColorearCeldaSeveridad()
    ' DECLARACIÓN DE VARIABLES
    Dim textoCelda As String
    Dim score As Double
    Dim bgColor As Long
    Dim fontColor As WdColor ' Usamos el tipo de dato específico para colores de Word
    Dim colorAplicar As Boolean
    Dim rng As Range

    ' 1. VERIFICAR QUE EL CURSOR ESTÁ EN UNA TABLA
    If Selection.Information(wdWithInTable) = False Then
        MsgBox "Por favor, coloque el cursor dentro de una celda de tabla.", vbInformation, "Operación Cancelada"
        Exit Sub
    End If

    ' 2. OBTENER Y LIMPIAR EL TEXTO DE LA CELDA
    ' Usamos el rango de la primera celda de la selección para ser más precisos
    Set rng = Selection.Cells(1).Range
    ' Excluimos el carácter de fin de celda invisible (muy importante)
    rng.End = rng.End - 1
    ' Limpiamos y convertimos a mayúsculas para que la comparación no falle
    textoCelda = UCase(Trim(rng.Text))
    
    ' Inicializamos la bandera
    colorAplicar = False

    ' 3. LÓGICA PARA DECIDIR EL COLOR BASADO EN EL CONTENIDO
    
    ' Primero, intentamos ver si el contenido es un número
    If IsNumeric(textoCelda) Then
        score = CDbl(textoCelda) ' Convertimos el texto a número
        
        If score >= 9# And score <= 10# Then
            bgColor = RGB(&H70, &H30, &HA0) ' Crítica (Púrpura)
            fontColor = wdColorWhite
            colorAplicar = True
        ElseIf score >= 7# And score < 9# Then
            bgColor = RGB(255, 0, 0) ' Alta (Rojo)
            fontColor = wdColorWhite
            colorAplicar = True
        ElseIf score >= 4# And score < 7# Then
            bgColor = RGB(255, 255, 0) ' Media (Amarillo)
            fontColor = wdColorBlack
            colorAplicar = True
        ElseIf score >= 0.1 And score < 4# Then
            bgColor = RGB(0, 176, 80) ' Baja (Verde)
            fontColor = wdColorWhite
            colorAplicar = True
        End If
    Else
        ' Si no es un número, lo tratamos como texto
        Select Case textoCelda
            Case "CRÍTICA", "CRITICAL"
                bgColor = RGB(&H70, &H30, &HA0) ' #7030A0
                fontColor = wdColorWhite
                colorAplicar = True
            Case "ALTA", "HIGH"
                bgColor = RGB(255, 0, 0) ' #FF0000
                fontColor = wdColorWhite
                colorAplicar = True
            Case "MEDIA", "MEDIUM"
                bgColor = RGB(255, 255, 0) ' #FFFF00
                fontColor = wdColorBlack
                colorAplicar = True
            Case "BAJA", "LOW"
                bgColor = RGB(0, 176, 80) ' #00B050
                fontColor = wdColorWhite
                colorAplicar = True
        End Select
    End If

    ' 4. APLICAR EL FORMATO DIRECTAMENTE A LA SELECCIÓN
    If colorAplicar Then
        ' Seleccionamos toda la celda para asegurar que el formato se aplique correctamente
        Selection.Cells(1).Select
        
        ' Aplicamos el formato de fondo usando tu método, que es el correcto.
        ' Esto asegura un color sólido y limpio.
        With Selection.Shading
            .Texture = wdTextureNone
            .ForegroundPatternColor = wdColorAutomatic
            .BackgroundPatternColor = bgColor
        End With
        
        ' Aplicamos el color de la fuente
        Selection.Font.color = fontColor
    End If

End Sub

Sub CYB_016_FormatearCodigoEnRango(ByVal selRange As Range)
    Dim colorKeyword As Long: colorKeyword = RGB(0, 0, 255)
    Dim colorCmdlet As Long: colorCmdlet = RGB(13, 35, 116)
    Dim colorString As Long: colorString = RGB(210, 50, 50)
    Dim colorComment As Long: colorComment = RGB(0, 128, 0)
    Dim colorType As Long: colorType = RGB(43, 145, 175)
    Dim colorOperator As Long: colorOperator = RGB(150, 150, 150)
    Dim colorNumber As Long: colorNumber = RGB(200, 100, 0)
    Dim colorDefault As Long: colorDefault = RGB(0, 0, 0)

    Dim keywords As Variant: keywords = Array("if", "then", "else", "elif", "fi", "case", "esac", "for", "in", "do", "done", _
                                              "while", "until", "function", "return", "select", "break", "continue", "true", "false", "null")

    Dim shellBuiltins As Variant: shellBuiltins = Array("echo", "read", "export", "local", "unset", "source", "alias", "bg", "fg", _
                                                         "cd", "pwd", "let", "eval", "exec", "exit", "history", "kill", _
                                                         "printf", "test", "trap")

    Dim psCmdlets As Variant: psCmdlets = Array("Write-Host", "Write-Output", "Write-Error", "Write-Warning", "Write-Verbose", _
                                                "Get-Help", "Get-Command", "Get-Process", "Get-Service", "Get-ChildItem", _
                                                "Get-Content", "Set-Content", "Add-Content", "Clear-Content", "Get-Location", _
                                                "Set-Location", "New-Item", "Remove-Item", "Copy-Item", "Move-Item", _
                                                "Rename-Item", "Invoke-Expression")

    Dim sqlKeywords As Variant: sqlKeywords = Array("SELECT", "FROM", "WHERE", "INSERT", "INTO", "VALUES", "UPDATE", "SET", "DELETE", _
                                                    "CREATE", "TABLE", "DATABASE", "ALTER", "DROP", "INDEX", "JOIN", "LEFT", "RIGHT", _
                                                    "INNER", "OUTER", "ON", "GROUP", "BY", "ORDER", "ASC", "DESC", "AS", "DISTINCT", _
                                                    "AND", "OR", "NOT")

    Dim commonTypes As Variant: commonTypes = Array("$true", "$false", "$null", "[string]", "[int]", "[bool]", "[array]", _
                                                    "[hashtable]", "[datetime]")

    Dim commonOperators As Variant: commonOperators = Array("-eq", "-ne", "-gt", "-ge", "-lt", "-le", "-and", "-or", "-not", _
                                                            "-like", "-match", "|", ">", "<", ">>")

    ' Aplicar colores
    CYB_018_FormatPattern selRange, """", """", colorString, False
    CYB_018_FormatPattern selRange, "'", "'", colorString, False
    CYB_018_FormatPattern selRange, "#", vbCr, colorComment, True
    CYB_018_FormatPattern selRange, "//", vbCr, colorComment, True
    CYB_019_FormatWithWildcards selRange, "[0-9]{1,}", colorNumber
    CYB_017_FormatKeywords selRange, commonTypes, colorType, True
    CYB_019_FormatWithWildcards selRange, "\$[a-zA-Z0-9_]{1,}", colorType
    CYB_017_FormatKeywords selRange, sqlKeywords, colorKeyword, True
    CYB_017_FormatKeywords selRange, psCmdlets, colorCmdlet, True
    CYB_017_FormatKeywords selRange, shellBuiltins, colorCmdlet, True
    CYB_017_FormatKeywords selRange, keywords, colorKeyword, True
    CYB_017_FormatKeywords selRange, commonOperators, colorOperator, False
End Sub

Private Sub CYB_017_FormatKeywords(ByVal rng As Range, ByVal keywords As Variant, ByVal color As Long, ByVal matchWord As Boolean)
    Dim keyword As Variant
    Dim findRange As Range
    Set findRange = rng.Duplicate

    For Each keyword In keywords
        With findRange.Find
            .ClearFormatting
            .Text = keyword
            .Forward = True
            .Wrap = wdFindStop
            .MatchCase = False
            .MatchWholeWord = matchWord
            .MatchWildcards = False

            Do While .Execute
                If findRange.InRange(rng) Then
                    findRange.Font.color = color
                End If
                findRange.Collapse wdCollapseEnd
            Loop
        End With
        Set findRange = rng.Duplicate
    Next keyword
End Sub

Private Sub CYB_018_FormatPattern(ByVal rng As Range, ByVal startChar As String, ByVal endChar As String, ByVal color As Long, ByVal toEndOfLine As Boolean)
    Dim findRange As Range
    Set findRange = rng.Duplicate

    With findRange.Find
        .ClearFormatting
        .Text = startChar
        .Forward = True
        .Wrap = wdFindStop
        .MatchCase = True

        Do While .Execute
            If findRange.InRange(rng) Then
                Dim endRange As Range
                Set endRange = findRange.Duplicate

                If toEndOfLine Then
                    endRange.End = endRange.Paragraphs(1).Range.End - 1
                Else
                    endRange.Collapse wdCollapseEnd
                    If Not endRange.Find.Execute(FindText:=endChar, Forward:=True) Then Exit Do
                End If

                Dim highlightRange As Range
                Set highlightRange = rng.Duplicate
                highlightRange.Start = findRange.Start
                highlightRange.End = endRange.End
                highlightRange.Font.color = color
            End If
            findRange.Collapse wdCollapseEnd
        Loop
    End With
End Sub

Private Sub CYB_019_FormatWithWildcards(ByVal rng As Range, ByVal pattern As String, ByVal color As Long)
    Dim findRange As Range
    Set findRange = rng.Duplicate

    With findRange.Find
        .ClearFormatting
        .Text = pattern
        .Forward = True
        .Wrap = wdFindStop
        .MatchWildcards = True

        Do While .Execute
            If findRange.InRange(rng) Then
                findRange.Font.color = color
            End If
            findRange.Collapse wdCollapseEnd
        Loop
    End With
End Sub




Sub CYB_020_FormatearCodigoHTMLSeleccionado()
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



Sub CYB_021_PalabrasClaveVerde_Corregida()

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
        regex.pattern = patterns(i) ' Establecer el patrón actual

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

Sub CYB_022_CensurarIPs_X_Dinamica_Segura()
    Dim i As Integer, j As Integer
    Dim strX3 As String, strX4 As String
    Dim rng As Range
    
    ' Desactivamos la actualización de pantalla para que sea rápido
    Application.ScreenUpdating = False
    
    ' Bucle para la longitud del 3er octeto (de 3 a 1 dígitos)
    For i = 3 To 1 Step -1
        ' Bucle para la longitud del 4to octeto (de 3 a 1 dígitos)
        For j = 3 To 1 Step -1
            
            ' Generamos la cadena de X correspondiente a la longitud (ej. XXX o XX)
            strX3 = String(i, "X")
            strX4 = String(j, "X")
            
            Set rng = ActiveDocument.Content
            
            With rng.Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .MatchWildcards = True
                
                ' Construimos el patrón exacto para esta combinación de longitudes
                ' <([0-9]{1,3})\.([0-9]{1,3})\.  -> Busca los dos primeros octetos (Grupo 1 y 2)
                ' ([0-9]{i})\.                   -> Busca el 3er octeto con longitud exacta 'i'
                ' ([0-9]{j})>                    -> Busca el 4to octeto con longitud exacta 'j'
                
                .Text = "<([0-9]{1,3})\.([0-9]{1,3})\.([0-9]{" & i & "})\.([0-9]{" & j & "})>"
                
                ' Reemplazo: Mantiene octetos 1 y 2, y pone las X calculadas
                .Replacement.Text = "\1.\2." & strX3 & "." & strX4
                
                ' Ejecuta el reemplazo en todo el documento de forma segura
                .Execute Replace:=wdReplaceAll
            End With
            
        Next j
    Next i
    
    ' Restauramos la pantalla
    Application.ScreenUpdating = True
    MsgBox "IPs censuradas respetando la cantidad de dígitos.", vbInformation

End Sub
