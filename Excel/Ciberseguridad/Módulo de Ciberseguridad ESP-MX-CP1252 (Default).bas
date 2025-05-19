Attribute VB_Name = "ExcelModuloCiberseguridad"




Sub ApplyParagraphFormattingToCell(cell As Object)
Dim cellRange As Object
Set cellRange = cell.Range
    
    With cellRange.ParagraphFormat
        .SpaceBefore = 12
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .OutlineLevel = 10
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = 0
        .CollapsedByDefault = False
    End With
    
    
    With cellRange.ParagraphFormat
        .SpaceBefore = 6
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .OutlineLevel = 10
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = 0
        .CollapsedByDefault = False
        .Alignment = 3
    End With
End Sub




Sub CYB008_LimpiarTextoYAgregarGuion()
    Dim celda As Range
    Dim Texto As String
    Dim textoLimpio As String
    Dim lineas As Variant
    Dim i As Integer
    Dim textoConGuiones As String
    Dim incluirGuion As Boolean
    
    
    For Each celda In Selection
        
        If Not IsEmpty(celda.value) Then
            Texto = celda.value
            
            
            lineas = Split(Texto, vbLf)
            textoLimpio = ""
            
            
            For i = LBound(lineas) To UBound(lineas)
                If Len(Trim(lineas(i))) > 0 Then
                    textoLimpio = textoLimpio & lineas(i) & vbLf
                End If
            Next i
            
            
            If Len(textoLimpio) > 0 Then
                textoLimpio = Left(textoLimpio, Len(textoLimpio) - 1)
            End If
            
            
            textoConGuiones = ""
            lineas = Split(textoLimpio, vbLf)
            incluirGuion = False
            
            
            For i = LBound(lineas) To UBound(lineas)
                If InStr(1, lineas(i), ":", vbTextCompare) > 0 And Not incluirGuion Then
                    
                    textoConGuiones = textoConGuiones & lineas(i) & vbLf
                    incluirGuion = True
                ElseIf incluirGuion Then
                    
                    If Len(Trim(lineas(i))) > 0 Then
                        textoConGuiones = textoConGuiones & " - " & lineas(i) & vbLf
                    Else
                        
                        textoConGuiones = textoConGuiones & vbLf
                    End If
                Else
                    
                    textoConGuiones = textoConGuiones & lineas(i) & vbLf
                End If
            Next i
            
            
            If Len(textoConGuiones) > 0 Then
                textoConGuiones = Left(textoConGuiones, Len(textoConGuiones) - 1)
            End If
            
            
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
    
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Selecciona la carpeta de salida"
        If .Show = -1 Then
            carpetaSalida = .SelectedItems(1)
        Else
            MsgBox "No se seleccion? ninguna carpeta.", vbExclamation
            Exit Sub
        End If
    End With
    
    
    Set ws = ActiveSheet
    If Not ws Is Nothing Then
        tempFileName = carpetaSalida & "\" & "SSIFO37-02_Matriz de seguimiento vulnerabilidades de aplicaciones.xlsx"
        ws.Copy
        Set wb = ActiveWorkbook
        
        With wb.Sheets(1)
            
            .Cells.Select
            Selection.RowHeight = 15
            
            
            Columns("A:A").Select
            With Selection
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
            End With
            
            
            Columns("C:C").Select
            With Selection
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
            
            
            .ListObjects(1).TableStyle = "TableStyleMedium1"
            
            
            Set tbl = .ListObjects(1)
            On Error Resume Next
            Set colSeveridad = tbl.ListColumns("Severidad")
            On Error GoTo 0
            
            
            If Not colSeveridad Is Nothing Then
                
                Set selectedRange = colSeveridad.DataBodyRange
                With selectedRange
                    
                    .FormatConditions.Add Type:=xlTextString, String:="CR�TICA", TextOperator:=xlContains
                    .FormatConditions(.FormatConditions.Count).SetFirstPriority
                    With .FormatConditions(1).Font
                        .Color = RGB(255, 255, 255)
                    End With
                    With .FormatConditions(1).Interior
                        .Color = RGB(112, 48, 160)
                    End With
                    
                    
                    .FormatConditions.Add Type:=xlTextString, String:="ALTA", TextOperator:=xlContains
                    .FormatConditions(.FormatConditions.Count).SetFirstPriority
                    With .FormatConditions(1).Font
                        .Color = RGB(255, 255, 255)
                    End With
                    With .FormatConditions(1).Interior
                        .Color = RGB(255, 0, 0)
                    End With
                    
                    
                    .FormatConditions.Add Type:=xlTextString, String:="MEDIA", TextOperator:=xlContains
                    .FormatConditions(.FormatConditions.Count).SetFirstPriority
                    With .FormatConditions(1).Font
                        .Color = RGB(0, 0, 0)
                    End With
                    With .FormatConditions(1).Interior
                        .Color = RGB(255, 255, 0)
                    End With
                    
                    
                    .FormatConditions.Add Type:=xlTextString, String:="BAJA", TextOperator:=xlContains
                    .FormatConditions(.FormatConditions.Count).SetFirstPriority
                    With .FormatConditions(1).Font
                        .Color = RGB(255, 255, 255)
                    End With
                    With .FormatConditions(1).Interior
                        .Color = RGB(0, 176, 80)
                    End With
                    
                    
                    .FormatConditions.Add Type:=xlTextString, String:="INFORMATIVA", TextOperator:=xlContains
                    .FormatConditions(.FormatConditions.Count).SetFirstPriority
                    With .FormatConditions(1).Font
                        .Color = RGB(0, 0, 0)
                    End With
                    With .FormatConditions(1).Interior
                        .Color = RGB(231, 230, 230)
                    End With
                End With
            Else
                MsgBox "No se encontr? la columna        "
            End If
        End With
        
        wb.Sheets(1).Name = "Vulnerabilidades"
        wb.Sheets(1).Range("A1").Select
        
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
    
    
    ws.Activate
    If Not ws Is Nothing Then
        tempFileName = carpetaSalida & "\" & fileName & ".xlsx"
        ws.Copy
        Set wb = ActiveWorkbook
        
        
        With wb.Sheets(1)
            
            .Cells.RowHeight = 15
            
            
            .Columns("A:A").HorizontalAlignment = xlCenter
            .Columns("A:A").VerticalAlignment = xlBottom
            
            
            .Columns("C:C").HorizontalAlignment = xlCenter
            .Columns("C:C").VerticalAlignment = xlCenter
            
            
            .ListObjects(1).TableStyle = "TableStyleMedium1"
            
            
            Set tbl = .ListObjects(1)
            On Error Resume Next
            Set colSeveridad = tbl.ListColumns("Severidad")
            On Error GoTo 0
            
            
            If Not colSeveridad Is Nothing Then
                
                Set selectedRange = colSeveridad.DataBodyRange
                With selectedRange
                    
                    .FormatConditions.Add Type:=xlTextString, String:="CR�TICA", TextOperator:=xlContains
                    .FormatConditions(.FormatConditions.Count).SetFirstPriority
                    With .FormatConditions(1).Font
                        .Color = RGB(255, 255, 255)
                    End With
                    With .FormatConditions(1).Interior
                        .Color = RGB(112, 48, 160)
                    End With
                    
                    
                    .FormatConditions.Add Type:=xlTextString, String:="ALTA", TextOperator:=xlContains
                    .FormatConditions(.FormatConditions.Count).SetFirstPriority
                    With .FormatConditions(1).Font
                        .Color = RGB(255, 255, 255)
                    End With
                    With .FormatConditions(1).Interior
                        .Color = RGB(255, 0, 0)
                    End With
                    
                    
                    .FormatConditions.Add Type:=xlTextString, String:="MEDIA", TextOperator:=xlContains
                    .FormatConditions(.FormatConditions.Count).SetFirstPriority
                    With .FormatConditions(1).Font
                        .Color = RGB(0, 0, 0)
                    End With
                    With .FormatConditions(1).Interior
                        .Color = RGB(255, 255, 0)
                    End With
                    
                    
                    .FormatConditions.Add Type:=xlTextString, String:="BAJA", TextOperator:=xlContains
                    .FormatConditions(.FormatConditions.Count).SetFirstPriority
                    With .FormatConditions(1).Font
                        .Color = RGB(255, 255, 255)
                    End With
                    With .FormatConditions(1).Interior
                        .Color = RGB(0, 176, 80)
                    End With
                    
                    
                    .FormatConditions.Add Type:=xlTextString, String:="INFORMATIVA", TextOperator:=xlContains
                    .FormatConditions(.FormatConditions.Count).SetFirstPriority
                    With .FormatConditions(1).Font
                        .Color = RGB(0, 0, 0)
                    End With
                    With .FormatConditions(1).Interior
                        .Color = RGB(231, 230, 230)
                    End With
                End With
            Else
                MsgBox "No se encontr? la columna        "
            End If
        End With
        
        wb.Sheets(1).Name = "Vulnerabilidades"
        wb.Sheets(1).Range("A1").Select
        
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
        valorActual = Trim(UCase(c.value))
        
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
            Case "9", "CR�TICA", "CRITICAL", "CR�TICO"
                c.value = "CR�TICA"
            Case "10", "CR�TICA", "CRITICAL", "CR�TICO"
                c.value = "CR�TICA"
                
            Case Else
                
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
    
    
    Set rng = Selection
    
    
    For Each cell In rng
        
        content = cell.value
        
        
        If content <> "" Then
            
            contentArray = Split(content, Chr(10))
            
            
            Set uniqueUrls = CreateObject("Scripting.Dictionary")
            
            
            For i = LBound(contentArray) To UBound(contentArray)
                If contentArray(i) <> "" Then
                    
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
            
            
            n = uniqueUrls.Count - 1
            ReDim uniqueArray(n)
            i = 0
            For Each key In uniqueUrls.Keys
                uniqueArray(i) = key
                i = i + 1
            Next
            
            
            For i = LBound(uniqueArray) To UBound(uniqueArray) - 1
                For j = i + 1 To UBound(uniqueArray)
                    If uniqueArray(i) > uniqueArray(j) Then
                        temp = uniqueArray(i)
                        uniqueArray(i) = uniqueArray(j)
                        uniqueArray(j) = temp
                    End If
                Next j
            Next i
            
            
            content = Join(uniqueArray, Chr(10))
            
            
            cell.value = content
        End If
    Next cell
End Sub

Sub CYB029_ReemplazarConURLs()
    Dim cell        As Range
    Dim parts       As Variant
    Dim url         As String
    Dim i           As Integer
    
    
    For Each cell In Selection
        If cell.value <> "" Then
            
            parts = Split(cell.value, ",")
            
            
            url = ""
            
            
            For i = LBound(parts) To UBound(parts)
                
                If InStr(parts(i), "|") > 0 Then
                    url = url & Mid(parts(i), InStr(parts(i), "|") + 1) & vbLf
                End If
            Next i
            
            
            If Len(url) > 0 Then
                url = Left(url, Len(url) - 1)
            End If
            
            
            url = Replace(url, """", "")
            
            
            cell.value = url
        End If
    Next cell
End Sub

Sub CYB037_AplicarFormatoCondicional()
    Dim selectedRange As Range
    
    
    On Error Resume Next
    Set selectedRange = Selection.SpecialCells(xlCellTypeConstants)
    On Error GoTo 0
    
    
    If selectedRange Is Nothing Then
        MsgBox "No hay celdas seleccionadas."
        Exit Sub
    End If
    
    
    With selectedRange
        .FormatConditions.Add Type:=xlTextString, String:="CR�TICA", TextOperator:=xlContains
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Font
            .Color = RGB(255, 255, 255)
        End With
        With .FormatConditions(1).Interior
            .Color = RGB(112, 48, 160)
        End With
        
        .FormatConditions.Add Type:=xlTextString, String:="ALTA", TextOperator:=xlContains
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Font
            .Color = RGB(255, 255, 255)
        End With
        With .FormatConditions(1).Interior
            .Color = RGB(255, 0, 0)
        End With
        
        .FormatConditions.Add Type:=xlTextString, String:="MEDIA", TextOperator:=xlContains
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Font
            .Color = RGB(0, 0, 0)
        End With
        With .FormatConditions(1).Interior
            .Color = RGB(255, 255, 0)
        End With
        
        .FormatConditions.Add Type:=xlTextString, String:="BAJA", TextOperator:=xlContains
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Font
            .Color = RGB(255, 255, 255)
        End With
        With .FormatConditions(1).Interior
            .Color = RGB(0, 176, 80)
        End With
        
        .FormatConditions.Add Type:=xlTextString, String:="INFORMATIVA", TextOperator:=xlContains
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Font
            .Color = RGB(0, 0, 0)
        End With
        With .FormatConditions(1).Interior
            .Color = RGB(231, 230, 230)
        End With
    End With
End Sub

Sub CYB033_ConvertirATextoEnOracion()
    Dim celda       As Range
    Dim Texto       As String
    Dim primeraLetra As String
    Dim restoTexto  As String
    
    
    For Each celda In Selection
        If Not IsEmpty(celda.value) Then
            Texto = celda.value
            
            Texto = LCase(Texto)
            
            primeraLetra = UCase(Left(Texto, 1))
            
            restoTexto = Mid(Texto, 2)
            
            celda.value = primeraLetra & restoTexto
        End If
    Next celda
End Sub

Sub CYB027_QuitarEspacios()
    Dim rng         As Range
    Dim c           As Range
    
    Set rng = Selection
    
    For Each c In rng
        c.value = Application.Trim(c.value)
    Next c
End Sub


Sub CYB009_ProcesadoCompletoSalidaHerramientas()
    Dim celda As Range
    Dim lineas As Variant
    Dim i As Integer

    ' Iterar sobre cada celda seleccionada
    For Each celda In Selection
        ' Verificar si la celda tiene texto
        If Not IsEmpty(celda.value) Then
            ' Reemplazar diferentes saltos de l�nea con vbLf
            Dim contenido As String
            contenido = Replace(Replace(Replace(celda.value, vbCrLf, vbLf), vbCr, vbLf), vbLf & vbLf, vbLf)
            
                     
            
            
            ' Si el contenido comienza con vbLf, quitarlo
            If Left(contenido, 1) = vbLf Then
                contenido = Mid(contenido, 2)
            End If
            
            ' Si el contenido termina con vbLf, quitarlo
            If Right(contenido, 1) = vbLf Then
                contenido = Left(contenido, Len(contenido) - 1)
            End If
            
            ' Dividir el contenido de la celda en un array de l�neas
            lineas = Split(contenido, vbLf)
            
            ' Crear un nuevo array para almacenar las l�neas no vac�as
            Dim lineasSinVacias() As String
            ReDim lineasSinVacias(0 To UBound(lineas))
            Dim idx As Integer
            idx = 0
            
            ' Iterar sobre cada l�nea del array
            For i = LBound(lineas) To UBound(lineas)
                ' Verificar si la l�nea est� vac�a y no agregarla al nuevo array
                If Trim(lineas(i)) <> "" Then
                    lineasSinVacias(idx) = lineas(i)
                    idx = idx + 1
                End If
            Next i
            
            ' Redimensionar el array resultante
            ReDim Preserve lineasSinVacias(0 To idx - 1)
            
            ' Unir el array de l�neas de nuevo en una cadena
            contenido = Join(lineasSinVacias, vbLf)

            ' Reemplazar "Nessus" por "The scanner tool"
            contenido = Replace(contenido, "Nessus", "The scanner tool")
            
            contenido = Replace(contenido, "/http", "/" & vbCrLf & "http")

            ' Asignar el contenido limpio a la celda
            celda.value = contenido
        End If
    Next celda

    MsgBox "Proceso completado: Se han eliminado los saltos de l�nea innecesarios y reemplazado 'Nessus' por 'The scanner tool'.", vbInformation
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

    
    htmlPattern = "<(\/?(p|a|ul|b|strong|i|u|br)[^>]*?)>|<\/p><p>"
    liPattern = "<li[^>]*?>"
    cleanHtmlPattern = "<[^>]+>"

    
    Set rng = Selection

    For Each celda In rng
        If Not IsEmpty(celda.value) Then
            If Not celda.HasFormula Then

                
                celda.value = Replace(celda.value, Chr(9), " ")
                
                celda.value = Application.Trim(celda.value)

                
                celda.value = RegExpReemplazar(celda.value, liPattern, vbLf)

                
                celda.value = RegExpReemplazar(celda.value, cleanHtmlPattern, vbNullString)

                
                lineas = Split(celda.value, vbLf)
                cleanOutput = ""
                lastLineWasEmpty = False

                
                i = LBound(lineas)
                Do While i <= UBound(lineas) And Trim(lineas(i)) = ""
                    i = i + 1
                Loop

                
                For i = i To UBound(lineas)
                    If Trim(lineas(i)) <> "" Then
                        cleanOutput = cleanOutput & lineas(i) & vbLf
                        lastLineWasEmpty = False
                    ElseIf Not lastLineWasEmpty Then
                        cleanOutput = cleanOutput & vbLf
                        lastLineWasEmpty = True
                    End If
                Next i

                
                If Len(cleanOutput) > 0 And Right(cleanOutput, 1) = vbLf Then
                    cleanOutput = Left(cleanOutput, Len(cleanOutput) - 1)
                End If

                
                Texto = ReemplazarEntidadesHtml(cleanOutput)

                
                NuevoTexto = Trim(Texto)

                
                celda.value = NuevoTexto
            End If
        End If
    Next celda

    MsgBox "Proceso completado: Se han eliminado los saltos de l�nea al inicio y limpiado el texto.", vbInformation
End Sub



Sub CYB010_AgregarSaltosLineaATextoGuiones()

    Dim celda As Range
    Dim Texto As String
    Dim partes() As String
    Dim i As Integer

    
    For Each celda In Selection
        
        If celda.HasFormula = False Then
            Texto = celda.value
            
            If InStr(Texto, "-") > 0 Then
                
                partes = Split(Texto, "-")
                
                
                Texto = partes(0)
                
                For i = 1 To UBound(partes)
                    
                    Texto = Texto & vbNewLine & "- " & partes(i)
                Next i
                
                
                celda.value = Texto
            End If
        End If
    Next celda
    
End Sub


Sub CYB011_BulletsAGuiones()

 Dim celda As Range
    For Each celda In Selection
        If Not IsEmpty(celda.value) Then
            celda.value = Replace(celda.value, "�", "-")
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
    
    
    For Each celda In Selection
        
        If Not IsEmpty(celda.value) Then
            
            resultado = ""
            
            
            contenido = Replace(celda.value, """", "")
            
            
            contenido = Replace(Replace(Replace(contenido, vbCrLf, vbLf), vbCr, vbLf), vbLf & vbLf, vbLf)
            
            
            If Left(contenido, 1) = vbLf Then
                contenido = Mid(contenido, 2)
            End If
            
            
            If Right(contenido, 1) = vbLf Then
                contenido = Left(contenido, Len(contenido) - 1)
            End If
            
            
            lineas = Split(contenido, vbLf)
            
            
            If UBound(lineas) >= 0 Then
                
                ReDim lineasSinVacias(0 To UBound(lineas))
                idx = 0
                
                
                For i = LBound(lineas) To UBound(lineas)
                    
                    If Trim(lineas(i)) <> "" Then
                        
                        startPos = InStr(1, lineas(i), "http")
                        Do While startPos > 0
                            
                            commaPos = InStr(startPos, lineas(i), ",")
                            spacePos = InStr(startPos, lineas(i), " ")
                            
                            If commaPos > 0 And (commaPos < spacePos Or spacePos = 0) Then
                                endPos = commaPos
                            ElseIf spacePos > 0 Then
                                endPos = spacePos
                            Else
                                endPos = Len(lineas(i)) + 1
                            End If
                            
                            
                            url = Mid(lineas(i), startPos, endPos - startPos)
                            
                            
                            resultado = resultado & url & vbCrLf
                            
                            
                            startPos = InStr(startPos + 1, lineas(i), "http")
                        Loop
                    End If
                Next i
                
                
                If Len(resultado) > 0 Then
                    If Right(resultado, 1) = vbCrLf Then
                        resultado = Left(resultado, Len(resultado) - 1)
                    End If
                End If
                
                
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
    
    
    Set objShell = CreateObject("WScript.Shell")
    
    
    For Each celda In Selection
        
        ip = Trim(celda.value)
        
        
        If ip <> "" Then
            respuesta = False
            
            
            For i = 1 To 3
                
                Set objExec = objShell.Exec("ping -n 1 -w 500 " & ip)
                resultado = objExec.StdOut.ReadAll
                
                
                If InStr(1, resultado, "TTL=", vbTextCompare) > 0 Then
                    respuesta = True
                    Exit For
                End If
            Next i
            
            
            If respuesta Then
                celda.Interior.Color = RGB(144, 238, 144)
            Else
                celda.Interior.Color = RGB(169, 169, 169)
            End If
        End If
    Next celda
    
    
    Set objShell = Nothing
    Set objExec = Nothing
    
    MsgBox "Ping completado.", vbInformation, "Finalizado"
End Sub


Sub CYB012_ObtenerIPs()
    Dim celda As Range
    Dim hostname As String
    Dim objShell As Object
    Dim objExec As Object
    Dim resultado As String
    Dim ipExtraida As String
    Dim listaIPs As String
    Dim lineas() As String
    Dim i As Integer
    Dim encontrado As Boolean
    
    Set objShell = CreateObject("WScript.Shell")
    listaIPs = ""
    
    ' Iteramos sobre las celdas seleccionadas
    For Each celda In Selection
        hostname = Trim(celda.value)
        
        If hostname <> "" Then
            ' Ejecutamos el comando nslookup
            Set objExec = objShell.Exec("nslookup " & hostname)
            resultado = objExec.StdOut.ReadAll
            
            ' Dividimos el resultado en l�neas
            lineas = Split(resultado, vbCrLf)
            
            ' Inicializamos la variable de control de IP v�lida
            encontrado = False
            ipExtraida = "Unknown"
            
            ' Iteramos sobre cada l�nea del resultado
            For i = 0 To UBound(lineas)
                ' Si la l�nea contiene "Address:", tratamos de extraer la IP
                If InStr(lineas(i), "Address:") > 0 Then
                    ipExtraida = Trim(Split(lineas(i), ":")(1))
                    
                    ' Si la IP no es la del servidor DNS (192.168.0.1) y es v�lida, la usamos
                    If ipExtraida <> "192.168.0.1" And ipExtraida <> "" Then
                        encontrado = True
                        Exit For
                    End If
                End If
            Next i
            
            ' Si no se encontr� una IP v�lida, se asigna "Unknown"
            If Not encontrado Then
                ipExtraida = "Unknown"
            End If
            
            ' A�adimos la IP (o "Unknown") a la lista
            listaIPs = listaIPs & ipExtraida & vbCrLf
        End If
    Next celda
    
    ' Copiamos las IPs al portapapeles si se encontraron
    If listaIPs <> "" Then
        CopiarAlPortapapeles listaIPs
    End If
    
    Set objShell = Nothing
    Set objExec = Nothing
    
    MsgBox "IPs obtenidas y copiadas al portapapeles.", vbInformation, "Finalizado"
End Sub

Sub CYB013_ReverseDNS()
    Dim celda As Range
    Dim ip As String
    Dim objShell As Object
    Dim objExec As Object
    Dim resultado As String
    Dim hostExtraido As String
    Dim lineas() As String
    Dim i As Integer
    Dim encontrado As Boolean
    Dim listaHostnames As String

    Set objShell = CreateObject("WScript.Shell")
    listaHostnames = ""

    For Each celda In Selection
        ip = Trim(celda.value)
        
        If ip <> "" Then
            Set objExec = objShell.Exec("nslookup " & ip)
            resultado = objExec.StdOut.ReadAll
            
            lineas = Split(resultado, vbCrLf)
            hostExtraido = "Unknown"
            encontrado = False
            
            For i = 0 To UBound(lineas)
                If InStr(lineas(i), "Name:") > 0 Then
                    hostExtraido = Trim(Split(lineas(i), ":")(1))
                    encontrado = True
                    Exit For
                End If
            Next i
            
            listaHostnames = listaHostnames & hostExtraido & vbCrLf
        End If
    Next celda

    If listaHostnames <> "" Then
        CopiarAlPortapapeles listaHostnames
        MsgBox "Hostnames copiados al portapapeles.", vbInformation, "Finalizado"
    Else
        MsgBox "No se encontraron resultados v�lidos.", vbExclamation, "Sin resultados"
    End If
    
    Set objShell = Nothing
    Set objExec = Nothing
End Sub

Sub CYB014_CheckHTTPHTTPS()
    Dim celda As Range
    Dim hostname As String
    Dim protocoloUsado As String
    Dim listaProtocolos As String
    Dim resultado As String

    On Error Resume Next

    For Each celda In Selection
        hostname = Trim(celda.value)
        protocoloUsado = "None"

        If hostname <> "" Then
            ' Probar HTTPS
            If URLResponde("https://" & hostname) Then
                protocoloUsado = "https"
            ElseIf URLResponde("http://" & hostname) Then
                protocoloUsado = "http"
            End If
        End If

        listaProtocolos = listaProtocolos & protocoloUsado & vbCrLf
    Next celda

    CopiarAlPortapapeles listaProtocolos
    MsgBox "Protocolos detectados y copiados al portapapeles.", vbInformation, "Finalizado"
End Sub

Function URLResponde(url As String) As Boolean
    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    
    On Error GoTo ErrorHandler
    
    ' Establecer tiempo de espera:
    '   - 5000 ms para conexi�n
    '   - 5000 ms para env�o
    '   - 5000 ms para recepci�n
    http.setTimeouts 5000, 5000, 5000, 5000
    
    http.Open "HEAD", url, False
    http.setRequestHeader "User-Agent", "Mozilla/5.0"
    http.Send
    
    If http.Status >= 200 And http.Status < 400 Then
        URLResponde = True
    Else
        URLResponde = False
    End If
    Exit Function

ErrorHandler:
    URLResponde = False
End Function



Sub CYB014_EscanearIPsDesdeSeleccionSinDuplicados()
    Dim celda As Range
    Dim ip As String
    Dim dictResultados As Object ' Dictionary para evitar escaneos duplicados
    Dim fs As Object
    Dim cmd As String, tempFile As String
    Dim resultado As String
    Dim fileNum As Integer, contenido As String
    Dim contador As Long
    Dim carpetaTemporal As String
    Dim nmapTimeoutSeconds As Long
    Dim startTime As Double, elapsed As Double
    Dim wsh As Object
    Dim nmapPath As String

    Set dictResultados = CreateObject("Scripting.Dictionary")
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set wsh = CreateObject("WScript.Shell")

    carpetaTemporal = fs.GetSpecialFolder(2) ' Carpeta temporal del sistema
    nmapTimeoutSeconds = 180 ' Tiempo m�ximo por escaneo (3 minutos aprox)
    contador = 0

    ' Verificar que Nmap est� disponible en el sistema
    On Error Resume Next
    nmapPath = wsh.Exec("cmd /c where nmap").StdOut.ReadLine
    On Error GoTo 0

    If nmapPath = "" Then
        MsgBox "Nmap no est� instalado o no est� en el PATH.", vbCritical
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    For Each celda In Selection
        ip = Trim(CStr(celda.value))
        If ip = "" Then GoTo Siguiente

        ' Si ya escaneamos esta IP, usamos el resultado guardado
        If dictResultados.exists(ip) Then
            celda.Offset(0, 1).value = dictResultados(ip)
            GoTo Siguiente
        End If

        ' Preparar archivo temporal �nico
        tempFile = carpetaTemporal & "\nmap_result_" & Replace(ip, ".", "_") & "_" & contador & ".txt"
        contador = contador + 1

        ' Ejecutar el escaneo con Nmap
        cmd = "cmd /c nmap --min-rate 5200 -Pn " & ip & " > """ & tempFile & """"
        wsh.Run cmd, 0, True ' Ejecuta el comando y espera a que termine
        startTime = Timer

        ' Esperar a que se genere el archivo o se agote el tiempo
        Do
            DoEvents
            If fs.fileExists(tempFile) Then
                If fs.GetFile(tempFile).Size > 0 Then Exit Do
            End If
            elapsed = Timer - startTime
            If elapsed < 0 Then elapsed = elapsed + 86400 ' Si cruzamos medianoche
            If elapsed > nmapTimeoutSeconds Then Exit Do
        Loop

        ' Leer el archivo o marcar timeout
        If fs.fileExists(tempFile) And fs.GetFile(tempFile).Size > 0 Then
            fileNum = FreeFile
            Open tempFile For Input As #fileNum
            contenido = Input$(LOF(fileNum), #fileNum)
            Close #fileNum
            resultado = contenido
        Else
            resultado = "Timeout o error al escanear IP: " & ip
        End If

        ' Guardar resultado en el diccionario y aplicarlo a la celda
        dictResultados.Add ip, resultado
        celda.Offset(0, 1).value = resultado

        ' Limpiar archivo temporal
        On Error Resume Next: Kill tempFile: On Error GoTo 0

Siguiente:
    Next celda

    Application.ScreenUpdating = True
    Application.EnableEvents = True

    MsgBox "Escaneo completado para IPs seleccionadas.", vbInformation
End Sub


Function RegExpReemplazar(ByVal text As String, ByVal replacePattern As String, ByVal replaceWith As String) As String
    
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
    
    text = Replace(text, "&lt;", "<")
    text = Replace(text, "&gt;", ">")
    text = Replace(text, "&amp;", "&")
    text = Replace(text, "&quot;", """")
    text = Replace(text, "&apos;", "
    text = Replace(text, "&#x27;", "
    text = Replace(text, "&#34;", """")
    text = Replace(text, "&#39;", "
    text = Replace(text, "&#160;", Chr(160))
    
    ReemplazarEntidadesHtml = text
End Function

Sub CYB026_OrdenaSegunColorRelleno()
    Dim celdaActual As Range
    Dim tabla As ListObject
    Dim ws As Worksheet
    Dim respuesta As VbMsgBoxResult
    Dim colores As Variant
    Dim i As Integer
    
    
    colores = Array(RGB(112, 48, 160), RGB(255, 0, 0), RGB(255, 255, 0), RGB(0, 176, 80))
    
    
    Set celdaActual = ActiveCell
    Set ws = ActiveSheet
    
    
    On Error Resume Next
    Set tabla = celdaActual.ListObject
    On Error GoTo 0
    
    If tabla Is Nothing Then
        MsgBox "Debes seleccionar una celda dentro de una tabla para ejecutar la ordenaci?n.", vbExclamation, "Error"
        Exit Sub
    End If
    
    
respuesta = MsgBox("Se ordenar� la tabla por la columna Orden: Morado ? Rojo ? Amarillo ? Verde." & vbCrLf & vbCrLf & _
                   "�Deseas continuar?", vbYesNo + vbQuestion, "Confirmaci�n")

    
    If respuesta <> vbYes Then Exit Sub
    
    
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

    salidaPruebaSeguridadKey = "�Salidas de herramienta�"
    metodoDeteccionKey = "�M�todo de detecci�n�"
    
    
    If replaceDic.exists(salidaPruebaSeguridadKey) And replaceDic.exists(metodoDeteccionKey) Then
        
        salidaPruebaSeguridadValue = CStr(replaceDic(salidaPruebaSeguridadKey))
        metodoDeteccionValue = CStr(replaceDic(metodoDeteccionKey))
        
        
        Set firstTable = WordDoc.Tables(1)
        numRows = firstTable.Rows.Count
        
        
        If Len(Trim(salidaPruebaSeguridadValue)) = 0 And Len(Trim(metodoDeteccionValue)) = 0 Then
            
            If numRows > 0 Then
                
                firstTable.Rows(numRows).Delete
                
                If numRows > 1 Then
                    firstTable.Rows(numRows - 1).Delete
                End If
            End If
      ElseIf Len(Trim(salidaPruebaSeguridadValue)) = 0 And firstTable.Tables.Count > 0 Then
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
    Set nuevoEstilo = docWord.Styles.Add(Name:=estilo, Type:=1)
    If Err.Number <> 0 Then
        MsgBox "No se pudo crear el estilo        "
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
    
    
    Set ws = ActiveSheet
    
    
    rutaBase = ws.Parent.path & "\"
    
    
    Set appWord = CreateObject("Word.Application")
    appWord.Visible = True
    Set docWord = appWord.Documents.Add
    
    
    On Error Resume Next
    Set tbl = ws.ListObjects("Tabla_pruebas_seguridad")
    On Error GoTo 0
    
    If tbl Is Nothing Then
        MsgBox "La tabla        "
        GoTo Cleanup
    End If
    
    
    For Each r In tbl.ListRows
        estilo = r.Range.Cells(1, tbl.ListColumns("Estilo").Index).value
        seccion = r.Range.Cells(1, tbl.ListColumns("Secci?n").Index).value
        descripcion = r.Range.Cells(1, tbl.ListColumns("Descripci?n").Index).value
        imagen = r.Range.Cells(1, tbl.ListColumns("Im?genes").Index).value
        parrafoResultados = r.Range.Cells(1, tbl.ListColumns("Resultado").Index).value
        
        
        imagenRutaCompleta = rutaBase & imagen
        
        
        If Not EstiloExiste(docWord, estilo) Then
            CrearEstilo docWord, estilo
        End If
        
        
        If Trim(seccion) <> "" Then
            With docWord.content.Paragraphs.Add
                .Range.text = seccion
                .Range.Style = docWord.Styles(estilo)
                .Range.InsertParagraphAfter
            End With
            
        End If
        
        
        If Trim(descripcion) <> "" Then
            With docWord.content.Paragraphs.Add
                .Range.text = descripcion
                .Range.Style = docWord.Styles("Normal")
                .Format.SpaceBefore = 12
            End With
        End If
        
        docWord.content.InsertParagraphAfter
        docWord.content.Paragraphs.Last.Range.Select
        
        
        If imagen <> "" Then
            
            If Dir(imagenRutaCompleta) <> "" Then
                
                docWord.content.InsertParagraphAfter
                Set rng = docWord.content.Paragraphs.Last.Range
                
                
                Set shape = docWord.InlineShapes.AddPicture(fileName:=imagenRutaCompleta, LinkToFile:=False, SaveWithDocument:=True, Range:=rng)
                
                
                shape.Range.ParagraphFormat.Alignment = 1
                
                
                docWord.content.InsertParagraphAfter
                Set rng = docWord.content.Paragraphs.Last.Range
                
                
                Set captionRange = rng.Duplicate
                captionRange.Select
                appWord.Selection.MoveLeft Unit:=1, Count:=1, Extend:=0
                appWord.CaptionLabels.Add Name:="Imagen"
                appWord.Selection.InsertCaption Label:="Imagen", TitleAutoText:="InsertarT�tulo1", _
                                                Title:="", Position:=1
                appWord.Selection.ParagraphFormat.Alignment = 1
                
                docWord.content.InsertAfter text:=" " & seccion
                
                
                docWord.content.InsertParagraphAfter
                
            Else
                MsgBox "La imagen        "
            End If
        End If
        
        
        If Trim(parrafoResultados) <> "" Then
            With docWord.content.Paragraphs.Add
                .Range.text = parrafoResultados
                .Range.Style = docWord.Styles("Normal")
                .Format.SpaceBefore = 12
            End With
        End If
    Next r
    
Cleanup:
    
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
    
    
    Set rng = Selection
    
    
    For Each cell In rng
        
        content = cell.value
        
        
        content = Replace(content, """", Chr(10))
        
        
        If content <> "" Then
            
            contentArray = Split(content, Chr(10))
            
            
            Set uniqueUrls = CreateObject("Scripting.Dictionary")
            
            
            For i = LBound(contentArray) To UBound(contentArray)
                If Trim(contentArray(i)) <> "" Then
                    
                    contentArray(i) = Trim(Replace(contentArray(i), Chr(13), ""))
                    contentArray(i) = Replace(contentArray(i), " ", "")
                    If InStr(1, contentArray(i), "wikipedia", vbTextCompare) = 0 Then
                        If Not uniqueUrls.exists(contentArray(i)) Then
                            uniqueUrls.Add contentArray(i), Nothing
                        End If
                    End If
                End If
            Next i
            
            
            n = uniqueUrls.Count - 1
            ReDim uniqueArray(n)
            i = 0
            For Each key In uniqueUrls.Keys
                uniqueArray(i) = key
                i = i + 1
            Next
            
            
            For i = LBound(uniqueArray) To UBound(uniqueArray) - 1
                For j = i + 1 To UBound(uniqueArray)
                    If uniqueArray(i) > uniqueArray(j) Then
                        temp = uniqueArray(i)
                        uniqueArray(i) = uniqueArray(j)
                        uniqueArray(j) = temp
                    End If
                Next j
            Next i
            
            
            newContent = Join(uniqueArray, Chr(10))
            
            
            newContent = Trim(newContent)
            
            
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
        
        keyValue = Split(line, ":")
        If UBound(keyValue) = 1 Then
            key = Trim(keyValue(0))
            value = Trim(Mid(keyValue(1), 2, Len(keyValue(1)) - 2))
            
            dataDict(key) = value
        End If
    Loop
    
    Close #fileNumber
End Sub

Sub CYB019_WordAppAlternativeReplaceParagraph(WordApp As Object, WordDoc As Object, wordToFind As String, replaceWord As String)
    Dim findInRange As Boolean
    Dim rng         As Object
    
    
    Set rng = WordDoc.content
    
    
    With rng.Find
        .text = wordToFind
        .Replacement.text = replaceWord
        .Forward = True
        .Wrap = 1
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    
    
    rng.Find.Execute Replace:=2
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
    
    
    If graficoIndex < 1 Or graficoIndex > WordDoc.InlineShapes.Count Then
        MsgBox "�ndice de gr?fico fuera de rango."
        ActualizarGraficoSegunDicionario = False
        Exit Function
    End If
    
    
    Set ils = WordDoc.InlineShapes(graficoIndex)
    
    If ils.Type = 12 And ils.HasChart Then
        Set Chart = ils.Chart
        If Not Chart Is Nothing Then
            
            Set ChartData = Chart.ChartData
            If Not ChartData Is Nothing Then
                ChartData.Activate
                Set ChartWorkbook = ChartData.Workbook
                If Not ChartWorkbook Is Nothing Then
                    Set SourceSheet = ChartWorkbook.Sheets(1)
                    
                    
                    lastRow = SourceSheet.Cells(SourceSheet.Rows.Count, 1).End(xlUp).row
                    If lastRow >= 2 Then
                        SourceSheet.Range("A2:B" & lastRow).ClearContents
                    End If
                    
                    
                    categoryRow = 2
                    For Each category In conteos.Keys
                        SourceSheet.Cells(categoryRow, 1).value = category
                        SourceSheet.Cells(categoryRow, 2).value = conteos(category)
                        categoryRow = categoryRow + 1
                    Next category
                    
                    
                    sheetIndex = 1
                    dataRangeAddress = CStr(ChartWorkbook.Sheets(sheetIndex).Name & "$A$1:$B$" & CStr(categoryRow - 1))
                    Debug.Print dataRangeAddress
                    
                    
                    On Error Resume Next
                    ChartWorkbook.Sheets(sheetIndex).ChartObjects(1).Chart.SetSourceData Source:=Range(dataRangeAddress)
                    If Err.Number <> 0 Then
                        MsgBox "Error al establecer el rango de datos: " & Err.Description
                        Err.Clear
                        ActualizarGraficoSegunDicionario = False
                        Exit Function
                    End If
                    On Error GoTo 0
                    
                    
                    On Error Resume Next
                    Chart.Refresh
                    If Err.Number <> 0 Then
                        MsgBox "Error al actualizar el gr?fico: " & Err.Description
                        Err.Clear
                        ActualizarGraficoSegunDicionario = False
                        Exit Function
                    End If
                    On Error GoTo 0
                    
                    
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
    Dim rng         As Range
    Dim tbl         As ListObject
    Dim WordApp     As Object
    Dim WordDoc     As Object
    Dim templatePath As String
    Dim outputPath  As String
    Dim replaceDic  As Object
    Dim cell        As Range
    Dim headerCell  As Range
    Dim colIndex    As Integer
    Dim explicacionTecnicaCol As Long
    Dim tipoTextoExplicacionCol As Long
    Dim rowCount    As Long
    Dim i           As Long
    Dim tempFolder  As String
    Dim tempFolderPath As String
    Dim saveFolder  As String
    Dim selectedRange As Range
    Dim documentsList() As String
    Dim fs          As Object
    Dim key         As Variant
    Dim explicacionTecnicaValue As String
    Dim tipoTextoValue As String
    Dim excelBasePath As String
    Dim finalDocumentPath As String
    Dim textoCelda  As String
    Dim cellRange   As Object
    
    On Error Resume Next
    Set selectedRange = Application.InputBox("Seleccione TODO el rango de la tabla (incluyendo encabezados) que contiene los datos", Type:=8)
    On Error GoTo 0
    
    If selectedRange Is Nothing Then
        MsgBox "No se ha seleccionado un rango v�lido.", vbExclamation
        Exit Sub
    End If
    
    On Error Resume Next
    Set tbl = selectedRange.ListObject
    On Error GoTo 0
    
    If tbl Is Nothing Then
        MsgBox "El rango seleccionado no est� dentro de una tabla (ListObject)." & vbCrLf & _
               "Aseg�rese de que los datos est�n formateados como tabla (Insertar > Tabla).", vbExclamation
        Exit Sub
    End If
    
    Set rng = tbl.Range
    
    templatePath = Application.GetOpenFilename("Documentos de Word (*.docx), *.docx", , "Seleccione un documento de Word como plantilla")
    If templatePath = "Falso" Then Exit Sub
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Seleccione la carpeta donde desea guardar los documentos generados"
        If .Show = -1 Then
            saveFolder = .SelectedItems(1)
            If Right(saveFolder, 1) <> "\" Then saveFolder = saveFolder & "\"
        Else
            Exit Sub
        End If
    End With
    
    On Error Resume Next
    Set WordApp = GetObject(, "Word.Application")
    If Err.Number <> 0 Then
        Set WordApp = CreateObject("Word.Application")
        Err.Clear
    End If
    On Error GoTo 0
    
    If WordApp Is Nothing Then
        MsgBox "No se pudo iniciar Microsoft Word.", vbCritical
        Exit Sub
    End If
    WordApp.Visible = True
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    tempFolder = Environ("TEMP") & "\tmp-" & Format(Now(), "yyyymmddhhmmss")
    If Not fs.FolderExists(tempFolder) Then MkDir tempFolder
    tempFolderPath = tempFolder & "\"
    
    If ThisWorkbook.path <> "" Then
        excelBasePath = ThisWorkbook.path & "\"
    Else
        MsgBox "Guarde primero el libro de Excel para poder resolver rutas relativas de im�genes.", vbExclamation
        
        WordApp.Quit
        Set WordApp = Nothing
        Set fs = Nothing
        Exit Sub
    End If
    
    explicacionTecnicaCol = 0
    tipoTextoExplicacionCol = 0
    For Each headerCell In tbl.HeaderRowRange.Cells
        Select Case Trim(headerCell.value)
            Case "Explicaci�n t�cnica"
                explicacionTecnicaCol = headerCell.Column - tbl.Range.Column + 1
            Case "Tipo de texto de explicaci�n t�cnica"
                tipoTextoExplicacionCol = headerCell.Column - tbl.Range.Column + 1
        End Select
    Next headerCell
    
    If explicacionTecnicaCol = 0 Then
        MsgBox "No se encontr� la columna"
        WordApp.Quit
        Set WordApp = Nothing
        Set fs = Nothing
        Exit Sub
    End If
    
    If tipoTextoExplicacionCol = 0 Then
        MsgBox "Advertencia: No se encontr� la columna, Se asumir� Texto Plano para todas las filas.", vbInformation
    End If
    
    rowCount = tbl.ListRows.Count
    ReDim documentsList(0 To rowCount - 1)
    
    For i = 1 To rowCount
        
        Set replaceDic = CreateObject("Scripting.Dictionary")
        
        For colIndex = 1 To tbl.ListColumns.Count
            Dim colName As String
            Dim cellValue As String
            colName = tbl.HeaderRowRange.Cells(1, colIndex).value
            cellValue = tbl.DataBodyRange.Cells(i, colIndex).value
            replaceDic("�" & colName & "�") = cellValue
            
            If colIndex = explicacionTecnicaCol Then
                explicacionTecnicaValue = cellValue
            End If
            If tipoTextoExplicacionCol > 0 And colIndex = tipoTextoExplicacionCol Then
                tipoTextoValue = Trim(LCase(cellValue))
            ElseIf tipoTextoExplicacionCol = 0 Then
                tipoTextoValue = "texto plano"
            End If
        Next colIndex
        
        Dim tempDocPath As String
        tempDocPath = tempFolderPath & "Documento_" & i & ".docx"
        fs.CopyFile templatePath, tempDocPath, True
        Set WordDoc = WordApp.Documents.Open(tempDocPath)
        WordDoc.Activate
        
        For Each key In replaceDic.Keys
            Dim placeholder As String
            Dim replacementValue As String
            placeholder = CStr(key)
            replacementValue = CStr(replaceDic(key))
            
            If placeholder = "�Explicaci�n t�cnica�" Then
                If tipoTextoValue = "markdown" Then
                    RawPrint explicacionTecnicaValue
                    'explicacionTecnicaValue = EliminarLineasVaciasdeString(explicacionTecnicaValue)
                    RawPrint explicacionTecnicaValue
                    InsertarTextoMarkdownEnWordConFormato WordApp, WordDoc, placeholder, explicacionTecnicaValue, excelBasePath, False, "Cuerpo de tabla"
                    SustituirTextoMarkdownPorImagenes WordApp, WordDoc, excelBasePath
                Else
                    WordAppReemplazarParrafo WordApp, WordDoc, placeholder, replacementValue
                End If
            ElseIf placeholder = "�Descripci�n�" Then
                WordAppReemplazarParrafo WordApp, WordDoc, placeholder, replacementValue
            ElseIf placeholder = "�Propuesta de remediaci�n�" Then
                WordAppReemplazarParrafo WordApp, WordDoc, placeholder, replacementValue
            ElseIf placeholder = "�Referencias�" Then
                WordAppReemplazarParrafo WordApp, WordDoc, placeholder, replacementValue
            Else
                WordAppReemplazarParrafo WordApp, WordDoc, placeholder, replacementValue
            End If
        Next key
        
        FormatearCeldaNivelRiesgo WordDoc.Tables(1).cell(1, 2)
        
        textoCelda = Trim(Replace(WordDoc.Tables(1).cell(3, 1).Range.text, Chr(13) & Chr(7), ""))
        If textoCelda = "AMENAZA" Then
            FormatearParrafosGuionesCelda WordDoc.Tables(1).cell(3, 2)
            AplicarNegritaPalabrasClaveEnCeldaWord WordDoc.Tables(1).cell(3, 2)
        End If
        
        textoCelda = Trim(Replace(WordDoc.Tables(1).cell(4, 1).Range.text, Chr(13) & Chr(7), ""))
        If textoCelda = "PROPUESTA DE REMEDIACI�N" Then
            FormatearParrafosGuionesCelda WordDoc.Tables(1).cell(4, 2)
            AplicarNegritaPalabrasClaveEnCeldaWord WordDoc.Tables(1).cell(4, 2)
        End If
        
        textoCelda = Trim(Replace(WordDoc.Tables(1).cell(5, 1).Range.text, Chr(13) & Chr(7), ""))
        If textoCelda = "AMENAZA" Or textoCelda = "PROPUESTA DE REMEDIACI�N" Then
            FormatearParrafosGuionesCelda WordDoc.Tables(1).cell(5, 2)
             AplicarNegritaPalabrasClaveEnCeldaWord WordDoc.Tables(1).cell(5, 2)
        End If
        
        textoCelda = Trim(Replace(WordDoc.Tables(1).cell(7, 1).Range.text, Chr(13) & Chr(7), ""))
        If textoCelda = "DETALLE DE PRUEBAS DE SEGURIDAD" Then
             With WordDoc.Tables(1).cell(8, 1).Range
                .Font.Color = wdColorBlack
               ' .ParagraphFormat.Alignment = wdAlignParagraphJustify
            End With
             
            'EliminarLineasVaciasEnCeldaTablaWord WordDoc, 1, 8, 1
            AjustarMarcadorCeldaEnTablaWord WordApp, WordDoc, 1, 8, 1
            EliminarUltimasFilasSiEsSalidaPruebaSeguridad WordDoc, replaceDic
            'AplicarFormatoCeldaEnTablaWord WordDoc, 1, 8, 1, "Cuerpo de tabla"
            AplicarNegritaPalabrasClaveEnCeldaWord WordDoc.Tables(1).cell(8, 1)
        End If
        
        finalDocumentPath = saveFolder & "Documento_Final_" & i & ".docx"
        WordDoc.SaveAs finalDocumentPath
        WordDoc.Close
        
        documentsList(i - 1) = finalDocumentPath
    Next i
    
    FusionarDocumentosInsertando WordApp, documentsList, saveFolder & "Documento_Completo_Fusionado.docx"
    
    WordApp.Quit
    Set WordApp = Nothing
    
    MsgBox "Se han generado los documentos de Word correctamente.", vbInformation
    
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
    
    
    campoArchivoPath = Application.GetOpenFilename("Archivos CSV (*.csv), *.csv", , "Seleccionar archivo CSV")
    If campoArchivoPath = "Falso" Then
        MsgBox "No se seleccion? ning�n archivo CSV. La macro se detendr?."
        Exit Sub
    End If
    
    
    Set replaceDic = CreateObject("Scripting.Dictionary")
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set ts = fileSystem.OpenTextFile(campoArchivoPath, 1, False, 0)
    
    
    Do Until ts.AtEndOfStream
        csvLine = ts.ReadLine
        partes = Split(csvLine, ",", 2)
        
        If UBound(partes) = 1 Then
            key = Trim(partes(0))
            value = Trim(partes(1))
            
            
            replaceDic(key) = value
        End If
    Loop
    ts.Close
    
    
    If replaceDic.exists("�Aplicaci?n�") Then
        appName = replaceDic("�Aplicaci?n�")
    Else
        MsgBox "No se encontr? el campo        "
        Exit Sub
    End If
    
    
    Set dlg = Application.FileDialog(msoFileDialogFilePicker)
    
    dlg.Title = "Seleccionar la plantilla de reporte t?cnico"
    dlg.Filters.Clear
    dlg.Filters.Add "Archivos de Word", "*.docx"
    If dlg.Show = -1 Then
        plantillaReportePath = dlg.SelectedItems(1)
    Else
        MsgBox "No se seleccion? ning�n archivo. La macro se detendr?."
        Exit Sub
    End If
    
    dlg.Title = "Seleccionar la plantilla de reporte ejecutivo"
    If dlg.Show = -1 Then
        plantillaReportePath2 = dlg.SelectedItems(1)
    Else
        MsgBox "No se seleccion? ning�n archivo. La macro se detendr?."
        Exit Sub
    End If
    
    dlg.Title = "Seleccionar la plantilla de tabla de vulnerabilidades"
    If dlg.Show = -1 Then
        plantillaVulnerabilidadesPath = dlg.SelectedItems(1)
    Else
        MsgBox "No se seleccion? ning�n archivo. La macro se detendr?."
        Exit Sub
    End If
    
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Seleccionar Carpeta de Salida"
        If .Show = -1 Then
            carpetaSalida = .SelectedItems(1)
        Else
            MsgBox "No se seleccion? ninguna carpeta. La macro se detendr?."
            Exit Sub
        End If
    End With
    
    
    carpetaSalida = carpetaSalida & "\AV " & appName
    On Error Resume Next
    MkDir carpetaSalida
    On Error GoTo 0
    
    
    On Error Resume Next
    Set WordApp = CreateObject("Word.Application")
    On Error GoTo 0
    
    If WordApp Is Nothing Then
        MsgBox "No se puede iniciar Microsoft Word."
        Exit Sub
    End If
    
    
    tempFolder = Environ("TEMP") & "\tmp-" & Format(Now(), "yyyymmddhhmmss")
    On Error Resume Next
    MkDir tempFolder
    On Error GoTo 0
    
    
    tempFolderGenerados = tempFolder & "\Documentos generados"
    On Error Resume Next
    MkDir tempFolderGenerados
    On Error GoTo 0
    
    
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
    
    
    On Error Resume Next
    Set selectedRange = Application.InputBox("Seleccione el rango de celdas que contienen las columnas a considerar", Type:=8)
    On Error GoTo 0
    
    If selectedRange Is Nothing Then
        MsgBox "No se ha seleccionado un rango v?lido.", vbExclamation
        Exit Sub
    End If
    
    Dim resultado As Boolean
    
    
    resultado = FunExportarHojaActivaAExcelINAI(carpetaSalida, appName)
    
    
    On Error Resume Next
    Set rng = selectedRange.ListObject.Range
    On Error GoTo 0
    
    If rng Is Nothing Then
        MsgBox "El rango seleccionado no est? dentro de una tabla.", vbExclamation
        Exit Sub
    End If
    
    
    Set severityCounts = CreateObject("Scripting.Dictionary")
    
    
    severidadColumna = -1
    For i = 1 To rng.Columns.Count
        If rng.Cells(1, i).value = "Severidad" Then
            severidadColumna = i
            Exit For
        End If
    Next i
    
    If severidadColumna = -1 Then
        MsgBox "No se encontr? la columna        "
        Exit Sub
    End If
    
    
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
    
    
    Set vulntypesCounts = CreateObject("Scripting.Dictionary")
    
    
    tiposvulnerabilidadColumna = -1
    For i = 1 To rng.Columns.Count
        If rng.Cells(1, i).value = "Tipo de vulnerabilidad" Then
            tiposvulnerabilidadColumna = i
            Exit For
        End If
    Next i
    
    If tiposvulnerabilidadColumna = -1 Then
        MsgBox "No se encontr? la columna        "
        Exit Sub
    End If
    
    
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
    
    
    countBAJA = IIf(severityCounts.exists("BAJA"), severityCounts("BAJA"), 0)
    countMEDIA = IIf(severityCounts.exists("MEDIA"), severityCounts("MEDIA"), 0)
    countALTA = IIf(severityCounts.exists("ALTA"), severityCounts("ALTA"), 0)
    countCRITICAS = IIf(severityCounts.exists("CR�TICOS"), severityCounts("CR�TICOS"), 0)
    
    
    totalVulnerabilidades = countBAJA + countMEDIA + countALTA + countCRITICAS
    
    
    tempDocVulnerabilidadesPath = tempFolder & "\Plantilla_Vulnerabilidades.docx"
    fileSystem.CopyFile plantillaVulnerabilidadesPath, tempDocVulnerabilidadesPath
    
    
    rowCount = rng.Rows.Count
    For i = 2 To rowCount
        Set replaceDic = CreateObject("Scripting.Dictionary")
        
        For Each cell In selectedRange.Rows(1).Cells
            replaceDic("�" & cell.value & "�") = rng.Cells(i, cell.Column).value
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
    
    
    finalDocumentPath = tempFolder & "\Tablas_vulnerabilidades.docx"
    FusionarDocumentosInsertando WordApp, documentsList, finalDocumentPath
    
    
    Set WordDoc = WordApp.Documents.Open(tempDocPath)
    secVulnerabilidades = "{{Secci?n de tablas de vulnerabilidades}}"
    
    Set rngReplace = WordDoc.content
    rngReplace.Find.ClearFormatting
    With rngReplace.Find
        .text = secVulnerabilidades
        .Replacement.text = ""
        .Forward = True
        .Wrap = 1
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
    
    
    Set rngReplace = WordDoc.content
    rngReplace.Find.ClearFormatting
    With rngReplace.Find
        .text = "�Total de vulnerabilidades�"
        .Replacement.text = totalVulnerabilidades
        .Forward = True
        .Wrap = 1
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
    
    
    FunActualizarGraficoSegunDicionario WordDoc, severityCounts, 1
    
    
    ActualizarGraficos WordDoc
    
    On Error Resume Next
    WordDoc.TablesOfContents(1).Update
    On Error GoTo 0
    
    
    WordDoc.SaveAs2 carpetaSalida & "\SSIFO14-03 Informe t?cnico.docx"
    
    
    nombrePDF = carpetaSalida & "\SSIFO14-03 Informe t?cnico.pdf"
    WordDoc.ExportAsFixedFormat OutputFileName:= _
                                nombrePDF, ExportFormat:= _
                                17, OpenAfterExport:=True, OptimizeFor:= _
                                0, Range:=0, From:=1, To:=1, _
                                Item:=0, IncludeDocProps:=True, KeepIRM:=True, _
                                CreateBookmarks:=1, DocStructureTags:=True, _
                                BitmapMissingFonts:=True, UseISO19005_1:=False
    WordDoc.Close False
    
    
    Set WordDoc = WordApp.Documents.Open(tempDocPath2)
    
    Set rngReplace = WordDoc.content
    rngReplace.Find.ClearFormatting
    With rngReplace.Find
        .text = "�Total de vulnerabilidades�"
        .Replacement.text = totalVulnerabilidades
        .Forward = True
        .Wrap = 1
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
    
    
    FunActualizarGraficoSegunDicionario WordDoc, severityCounts, 1
    
    
    FunActualizarGraficoSegunDicionario WordDoc, vulntypesCounts, 2
    
    ActualizarGraficos WordDoc
    
    
    On Error Resume Next
    WordDoc.TablesOfContents(1).Update
    On Error GoTo 0
    
    
    WordDoc.SaveAs2 carpetaSalida & "\SSIFO15-03 Informe Ejecutivo.docx"
    
    nombrePDF = carpetaSalida & "\SSIFO14-03 Informe Ejecutivo.pdf"
    WordDoc.ExportAsFixedFormat OutputFileName:= _
                                nombrePDF, ExportFormat:= _
                                17, OpenAfterExport:=True, OptimizeFor:= _
                                0, Range:=0, From:=1, To:=1, _
                                Item:=0, IncludeDocProps:=True, KeepIRM:=True, _
                                CreateBookmarks:=1, DocStructureTags:=True, _
                                BitmapMissingFonts:=True, UseISO19005_1:=False
    WordDoc.Close False
    
    
    WordApp.Quit
    Set WordDoc = Nothing
    Set WordApp = Nothing
    Set fileSystem = Nothing
    
    
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


    
    Set replaceDic = CreateObject("Scripting.Dictionary")
    
    
    On Error Resume Next
    Set selectedRange = Application.InputBox("Seleccione una fila con los valores para procesar", Type:=8)
    On Error GoTo 0
    
    If selectedRange Is Nothing Then
        MsgBox "No se ha seleccionado un rango v?lido.", vbExclamation
        Exit Sub
    End If
    
    
    If Not selectedRange.ListObject Is Nothing Then
        
        Set tableRange = selectedRange.ListObject.Range
    Else
        
        MsgBox "El rango seleccionado no est? dentro de una tabla.", vbExclamation
        Exit Sub
    End If
    
    
    For Each cell In tableRange.Rows(1).Cells
        If cell.value <> "" Then
            
            Set headerRow = tableRange.Rows(1)
            Exit For
        End If
    Next cell
    
    If headerRow Is Nothing Then
        MsgBox "No se han encontrado encabezados en el rango seleccionado.", vbExclamation
        Exit Sub
    End If
    
    
    For i = 1 To selectedRange.Columns.Count
        key = "�" & headerRow.Cells(1, i).value & "�"
        
        
        If replaceDic.exists(key) Then
            MsgBox "Se ha encontrado un encabezado duplicado: " & headerRow.Cells(1, i).value & _
                   vbCrLf & "Por favor, corrige los encabezados duplicados y vuelve a ejecutar la macro.", vbExclamation
            Exit Sub
        End If
        
        
        value = selectedRange.Cells(1, i).value
        replaceDic.Add key, value
    Next i
    
    
    If replaceDic.exists("�Nombre de carpeta�") Then
        folderName = replaceDic("�Nombre de carpeta�")
    Else
        MsgBox "No se encontr? el campo        "
        Exit Sub
    End If
    
    
    carpetaSalida = carpetaSalida & "\" & folderName
    On Error Resume Next
    MkDir carpetaSalida
    On Error GoTo 0
    
    If replaceDic.exists("�Tipo de reporte�") Then
        Select Case replaceDic("�Tipo de reporte�")
            Case "T?cnico"
                
                
                If replaceDic.exists("�Ruta de la plantilla�") Then
                    plantillaReportePath = replaceDic("�Ruta de la plantilla�")
                Else
                    MsgBox "No se encontr? el campo        "
                    Exit Sub
                End If
                
                
                If Len(Dir(plantillaReportePath)) = 0 Then
                    MsgBox "La ruta de la plantilla no es v?lida o el archivo no existe: " & plantillaReportePath, vbExclamation
                    Exit Sub
                End If
                
                Set dlg = Application.FileDialog(msoFileDialogFilePicker)
                
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
                dlg.Filters.Clear
                dlg.Filters.Add "Archivos de Word", "*.docx; *.doc; *.dotx; *.dot"

                dlg.Title = "Seleccionar la plantilla de tabla de vulnerabilidades"
                If dlg.Show = -1 Then
                    plantillaVulnerabilidadesPath = dlg.SelectedItems(1)
                Else
                    MsgBox "No se seleccion? ning�n archivo. La macro se detendr?."
                    Exit Sub
                End If
                
                
                On Error Resume Next
                Set WordApp = CreateObject("Word.Application")
                On Error GoTo 0
                If WordApp Is Nothing Then
                    MsgBox "No se puede iniciar Microsoft Word."
                    Exit Sub
                End If
                
                
                tempFolder = Environ("TEMP") & "\tmp-" & Format(Now(), "yyyymmddhhmmss")
                On Error Resume Next
                MkDir tempFolder
                On Error GoTo 0
                
                
                tempFolderGenerados = tempFolder & "\Documentos_generados"
                On Error Resume Next
                MkDir tempFolderGenerados
                On Error GoTo 0
                
                
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
                
                
                On Error Resume Next
                Set selectedRange = Application.InputBox("Seleccione el rango de celdas que contienen las columnas a considerar", Type:=8)
                On Error GoTo 0
                
                If selectedRange Is Nothing Then
                    MsgBox "No se ha seleccionado un rango v?lido.", vbExclamation
                    Exit Sub
                End If
                
                
                If Not selectedRange.ListObject Is Nothing Then
                    
                    Set tableRange = selectedRange.ListObject.Range
                Else
                    
                    MsgBox "El rango seleccionado no est? dentro de una tabla.", vbExclamation
                    Exit Sub
                End If
                
                Dim resultado As Boolean
                
                
                resultado = FunExportarHojaActivaAExcelINAI(carpetaSalida, folderName, tableRange.Worksheet, replaceDic("�Nombre del reporte�"))
                
            Case "Tablas de vulnerabilidades"
                
                GenerarDocumentosVulnerabilidiadesWord (replaceDic("�Nombre del reporte�"))
                
            Case Else
                MsgBox "El tipo de reporte no es reconocido.", vbExclamation
                Exit Sub
        End Select
    Else
        MsgBox "No se encontr? el campo        "
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
    
    Set objWMI = GetObject("winmgmts:\\.\root\CIMV2")
    
    Set objProcesses = objWMI.ExecQuery("SELECT * FROM Win32_Process WHERE Name =
    
    
    For Each objProcess In objProcesses
        objProcess.Terminate
    Next
    
    
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
    
    
    Set WordApp = WordDoc.Application
    
    
    Set docContent = WordDoc.content
    
    
    For Each key In replaceDic.Keys
        
        With WordApp.Selection.Find
            .ClearFormatting
            .text = key
            .Forward = True
            .Wrap = 1
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            
            
            findInRange = .Execute
            Do While findInRange
                
                WordApp.Selection.text = CStr(replaceDic(key))
                
                findInRange = .Execute
            Loop
        End With
    Next key
    
    
    Set docContent = Nothing
    Set WordApp = Nothing
End Sub


Sub ActualizarGraficos(ByRef WordDoc As Object)
    
    On Error Resume Next
    
    
    Dim i           As Integer
    Dim Chart       As Object
    Dim ChartData   As Object
    Dim ChartWorkbook As Object
    
    For i = 1 To WordDoc.InlineShapes.Count
        With WordDoc.InlineShapes(i)
            
            If .Type = 12 And .HasChart Then
                Set Chart = .Chart
                If Not Chart Is Nothing Then
                    
                    Set ChartData = Chart.ChartData
                    If Not ChartData Is Nothing Then
                        ChartData.Activate
                        Set ChartWorkbook = ChartData.Workbook
                        If Not ChartWorkbook Is Nothing Then
                            
                            ChartWorkbook.Application.Visible = False
                            
                            ChartWorkbook.Close SaveChanges:=False
                        End If
                        
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
    Dim selectedRange As Range
    Dim documentsList() As String
    
    
    On Error Resume Next
    Set selectedRange = Application.InputBox("Seleccione el rango de celdas que contienen las columnas a considerar", Type:=8)
    On Error GoTo 0
    
    If selectedRange Is Nothing Then
        MsgBox "No se ha seleccionado un rango v?lido.", vbExclamation
        Exit Function
    End If
    
    
    On Error Resume Next
    Set rng = selectedRange.ListObject.Range
    On Error GoTo 0
    
    If rng Is Nothing Then
        MsgBox "El rango seleccionado no est? dentro de una tabla.", vbExclamation
        Exit Function
    End If
    
    
    templatePath = Application.GetOpenFilename("Documentos de Word (*.docx), *.docx", , "Seleccione un documento de Word como plantilla")
    If templatePath = "Falso" Then Exit Function
    
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Seleccione la carpeta donde desea guardar los documentos generados"
        If .Show = -1 Then
            saveFolder = .SelectedItems(1)
        Else
            Exit Function
        End If
    End With
    
    
    Set WordApp = CreateObject("Word.Application")
    WordApp.Visible = True
    
    
    Set WordDoc = WordApp.Documents.Open(templatePath)
    
    
    Set replaceDic = CreateObject("Scripting.Dictionary")
    
    
    rowCount = rng.Rows.Count
    For Each cell In selectedRange.Rows(1).Cells
        replaceDic("�" & cell.value & "�") = ""
    Next cell
    
    
    tempFolder = Environ("TEMP") & "\tmp-" & Format(Now(), "yyyymmddhhmmss")
    MkDir tempFolder
    
    
    Dim fs          As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    
    For i = 2 To rowCount
        
        Set replaceDic = CreateObject("Scripting.Dictionary")
        
        
        For Each cell In selectedRange.Rows(1).Cells
            replaceDic("�" & cell.value & "�") = rng.Cells(i, cell.Column).value
        Next cell
        
        
        fs.CopyFile templatePath, tempFolder & "\Tabla_" & i & ".docx"
        
        Set WordDoc = WordApp.Documents.Open(tempFolder & "\Tabla_" & i & ".docx")
        
        For Each key In replaceDic.Keys
            
            Debug.Print CStr(key)
            If CStr(key) = "�Descripci?n�" Then
                
                replaceDic(key) = TransformarTexto(replaceDic(key))
            End If
            If CStr(key) = "�Propuesta de remediaci?n�" Then
                replaceDic(key) = TransformarTexto(replaceDic(key))
            End If
            If CStr(key) = "Referencias" Then
                replaceDic(key) = TransformarTexto(replaceDic(key))
            End If
            
            WordAppReemplazarParrafo WordApp, WordDoc, CStr(key), CStr(replaceDic(key))
            
           
        Next key
        FormatearCeldaNivelRiesgo WordDoc.Tables(1).cell(1, 2)
        
        
        
        WordDoc.Save
        WordDoc.Close
        
        
        ReDim Preserve documentsList(i - 2)
        documentsList(i - 2) = tempFolder & "\Tabla_" & i & ".docx"
    Next i
    
    
    Dim finalDocumentPath As String
    finalDocumentPath = saveFolder & "\" & fileName & ".docx"
    FusionarDocumentosInsertando WordApp, documentsList, finalDocumentPath
    
    
    fs.MoveFolder tempFolder, saveFolder & "\Documentos_generados"
    
    
    WordApp.Quit
    Set WordApp = Nothing
    
    
    MsgBox "Se han generado los documentos de Word correctamente.", vbInformation
End Function



Sub FormatearParrafosGuionesCelda(cell As Object)
    Dim cellText As String
    Dim p As Object
    Dim rng As Object
    Dim posDosPuntos As Integer
    Dim strTexto As String

    
    cellText = Trim(Replace(cell.Range.text, vbCrLf, ""))
    cellText = Trim(Replace(cellText, vbCr, ""))
    cellText = Trim(Replace(cellText, vbLf, ""))
    cellText = Trim(Replace(cellText, Chr(7), ""))

    
    For Each p In cell.Range.Paragraphs
        strTexto = p.Range.text
        
        
        If Left(Trim(strTexto), 2) = "- " Then
            posDosPuntos = InStr(strTexto, ":")
            
            
            If posDosPuntos > 0 Then
                
                Set rng = p.Range
                rng.Start = p.Range.Start
                rng.End = p.Range.Start + posDosPuntos - 1
                rng.Font.Bold = True
                
                
                Set rng = p.Range
                rng.Start = p.Range.Start + posDosPuntos
                rng.End = p.Range.End
                rng.Font.Bold = False
            End If
        End If
    Next p
End Sub

Sub AplicarNegritaPalabrasClaveEnCeldaWord(ByVal celda As Object)
    Dim palabrasClaveParte1 As Variant
    Dim palabrasClaveParte2 As Variant
    Dim palabrasClaveParte3 As Variant
    Dim palabra As Variant
    Dim rng As Object
    Dim findObj As Object

    ' Dividiendo las palabras clave en tres partes
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

    ' Crear rango de la celda para b�squeda
    Set rng = celda.Range
    Set findObj = rng.Find
    
    ' Aplicar negrita a las palabras clave de cada parte
    For Each palabra In palabrasClaveParte1
        With findObj
            .ClearFormatting
            .text = palabra
            .Replacement.ClearFormatting
            .Replacement.Font.Bold = True
            .Forward = True
            .Wrap = 1 ' wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = True
            .MatchWildcards = False
            .Execute Replace:=2 ' wdReplaceAll
        End With
    Next palabra
    
    For Each palabra In palabrasClaveParte2
        With findObj
            .ClearFormatting
            .text = palabra
            .Replacement.ClearFormatting
            .Replacement.Font.Bold = True
            .Forward = True
            .Wrap = 1 ' wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = True
            .MatchWildcards = False
            .Execute Replace:=2 ' wdReplaceAll
        End With
    Next palabra
    
    For Each palabra In palabrasClaveParte3
        With findObj
            .ClearFormatting
            .text = palabra
            .Replacement.ClearFormatting
            .Replacement.Font.Bold = True
            .Forward = True
            .Wrap = 1 ' wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = True
            .MatchWildcards = False
            .Execute Replace:=2 ' wdReplaceAll
        End With
    Next palabra
End Sub



Sub FormatearCeldaNivelRiesgo(cell As Object)
    Dim cellText As String
    Dim cvssScore As Double
    Dim isNumber As Boolean
    
    
    cellText = Trim(Replace(cell.Range.text, vbCrLf, ""))
    cellText = Trim(Replace(cellText, vbCr, ""))
    cellText = Trim(Replace(cellText, vbLf, ""))
    cellText = Trim(Replace(cellText, Chr(7), ""))
    
    
    On Error Resume Next
    cvssScore = CDbl(cellText)
    isNumber = (Err.Number = 0)
    On Error GoTo 0
    
    
    If isNumber Then
        Select Case cvssScore
            Case Is >= 9
                cell.Shading.BackgroundPatternColor = 10498160
                cell.Range.Font.Color = 16777215
            Case Is >= 7
                cell.Shading.BackgroundPatternColor = 255
                cell.Range.Font.Color = 16777215
            Case Is >= 4
                cell.Shading.BackgroundPatternColor = 65535
                cell.Range.Font.Color = 0
            Case Is >= 0.1
                cell.Shading.BackgroundPatternColor = 5287936
                cell.Range.Font.Color = 16777215
            Case Else
                cell.Shading.BackgroundPatternColor = wdColorAutomatic
        End Select
    Else
        
        Select Case UCase(cellText)
            Case "CR�TICA"
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
    
    
    With regex
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .pattern = "([^.()\r\n-])(?![^(]*\)|[-])[^\S\r\n]*[\r\n]+"
    End With
    
    
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
    
    
    If graficoIndex < 1 Or graficoIndex > WordDoc.InlineShapes.Count Then
        MsgBox "�ndice de gr?fico fuera de rango."
        FunActualizarGraficoSegunDicionario = False
        Exit Function
    End If
    
    
    Set ils = WordDoc.InlineShapes(graficoIndex)
    
    If ils.Type = 12 And ils.HasChart Then
        Set Chart = ils.Chart
        If Not Chart Is Nothing Then
            
            Set ChartData = Chart.ChartData
            If Not ChartData Is Nothing Then
                ChartData.Activate
                Set ChartWorkbook = ChartData.Workbook
                If Not ChartWorkbook Is Nothing Then
                    Set SourceSheet = ChartWorkbook.Sheets(1)
                    
                    
                    lastRow = SourceSheet.Cells(SourceSheet.Rows.Count, 1).End(xlUp).row
                    If lastRow >= 2 Then
                        SourceSheet.Range("A2:B" & lastRow).ClearContents
                    End If
                    
                    
                    categoryRow = 2
                    For Each category In conteos.Keys
                        SourceSheet.Cells(categoryRow, 1).value = category
                        SourceSheet.Cells(categoryRow, 2).value = conteos(category)
                        categoryRow = categoryRow + 1
                    Next category
                    
                    
                    dataRangeAddress =
                    Debug.Print dataRangeAddress
                    
                    
                    On Error Resume Next
                    Set DataTable = SourceSheet.ListObjects(tableIndex)
                    On Error GoTo 0
                    
                    
                    If Not DataTable Is Nothing Then
                        
                        DataTable.Resize SourceSheet.Range("A1:B" & (categoryRow - 1))
                    Else
                        MsgBox "La tabla en el �ndice " & tableIndex & " no se encontr? en la hoja."
                    End If
                    
                    WordDoc.InlineShapes(graficoIndex).Chart.SetSourceData Source:=dataRangeAddress
                    
                    
                    Chart.Refresh
                    
                    
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
    
    
    Set searchRange = WordDoc.content
    
    
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
    
    
    Do While searchRange.Find.Execute
        
        searchRange.text = replaceWord
        
        
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

    
    Set celda = Selection

    
    If IsEmpty(celda) Then
        MsgBox "Seleccione una celda con un rango de IPs.", vbExclamation, "Error"
        Exit Sub
    End If

    ipRango = Trim(celda.value)

    
    partes = Split(ipRango, "-")
    
    
    If UBound(partes) <> 1 Then
        MsgBox "Formato inv?lido. Use: 10.0.1.60-10.0.1.78", vbExclamation, "Error"
        Exit Sub
    End If

    ipInicio = partes(0)
    ipFin = partes(1)

    
    numInicio = CInt(Split(ipInicio, ".")(3))
    numFin = CInt(Split(ipFin, ".")(3))

    
    If numInicio > numFin Then
        MsgBox "El rango de IPs es inv?lido.", vbExclamation, "Error"
        Exit Sub
    End If

    
    Dim baseIP As String
    baseIP = Left(ipInicio, InStrRev(ipInicio, "."))

    
    filaActual = celda.row + 1

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
    
    
    Set wb = ThisWorkbook
    Set ws = wb.ActiveSheet
    
    
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
    
    
    If Not celdaEnTabla Then
        MsgBox "La celda seleccionada no est? dentro de una tabla.", vbExclamation
        Exit Sub
    End If
    
    
    archivos = Application.GetOpenFilename("Archivos CSV (*.csv), *.csv", MultiSelect:=True, Title:="Seleccionar archivos CSV")
    If IsArray(archivos) = False Then Exit Sub
    
    
    For i = LBound(archivos) To UBound(archivos)
        
        Set wbCSV = Workbooks.Open(fileName:=archivos(i), Local:=True)
        Set wsCSV = wbCSV.Sheets(1)
        
        
        encabezados = wsCSV.UsedRange.Rows(1).value
        csvData = wsCSV.UsedRange.Offset(1, 0).value
        
        
        wbCSV.Close False
        
        
        Dim j As Integer
        For j = 1 To UBound(csvData, 1)
            Set fila = tbl.ListRows.Add
            
            For colCSVIndex = 1 To UBound(encabezados, 2)
                
                Select Case encabezados(1, colCSVIndex)
                    Case "Host": columnaCorrespondiente = "Identificador de detecci�n usado"
                    Case "CVE": columnaCorrespondiente = "CVE"
                    Case "CVSS v3.0 Base Score": columnaCorrespondiente = "CVSSScore"
                    Case "Metasploit": columnaCorrespondiente = "Exploits p�blicos"
                    Case "Plugin ID": columnaCorrespondiente = "Identificador original de la vulnerabilidad"
                    Case "Name": columnaCorrespondiente = "Nombre original de la vulnerabilidad"
                    Case "Protocol": columnaCorrespondiente = "Protocolo de transporte"
                    Case "Port": columnaCorrespondiente = "Puerto"
                    Case "See Also": columnaCorrespondiente = "Referencias"
                    Case "Plugin Output": columnaCorrespondiente = "Salidas de herramienta"
                    Case Else: columnaCorrespondiente = ""
                End Select
                
                
                If columnaCorrespondiente <> "" Then
                    columnaDestino = tbl.ListColumns(columnaCorrespondiente).Index
                    fila.Range(1, columnaDestino).value = csvData(j, colCSVIndex)
                End If
            Next colCSVIndex
            
            
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
    
    
    Set wb = ThisWorkbook
    Set ws = wb.ActiveSheet
    
    
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
    
    
    If Not celdaEnTabla Then
        MsgBox "La celda seleccionada no est? dentro de una tabla.", vbExclamation
        Exit Sub
    End If
    
    
    mensaje = "�Est? seguro que desea cargar datos de los archivos CSV en la tabla "
    respuesta = MsgBox(mensaje, vbYesNo + vbQuestion, "Confirmaci?n")
    If respuesta = vbNo Then Exit Sub
    
    
    archivos = Application.GetOpenFilename("Archivos CSV (*.csv), *.csv", MultiSelect:=True, Title:="Seleccionar archivos CSV")
    If Not IsArray(archivos) Then Exit Sub
    
    
    For Each archivo In archivos
        
        If LCase(Trim(Right(archivo, 4))) <> ".csv" Then
            MsgBox "El archivo " & archivo & " no es un archivo CSV.", vbExclamation
            Exit Sub
        End If
        
        
        Set wbCSV = Workbooks.Open(fileName:=archivo, Local:=True)
        Set wsCSV = wbCSV.Sheets(1)
        
        
        encabezados = wsCSV.UsedRange.Rows(1).value
        
        
        csvData = wsCSV.UsedRange.Offset(1, 0).value
        
        
        wbCSV.Close False
        
        
        For i = 1 To UBound(csvData, 1)
            Set fila = tbl.ListRows.Add
            
            For colCSVIndex = 1 To UBound(encabezados, 2)
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
                
                
                If columnaCorrespondiente <> "" Then
                    columnaDestino = tbl.ListColumns(columnaCorrespondiente).Index
                    fila.Range(1, columnaDestino).value = csvData(i, colCSVIndex)
                End If
            Next colCSVIndex
            
            
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

    
    Set wb = ThisWorkbook
    Set ws = wb.ActiveSheet
    
    
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
    
    
    mensaje = "�Est? seguro que desea cargar datos del archivo XML en la tabla "
    respuesta = MsgBox(mensaje, vbYesNo + vbQuestion, "Confirmaci?n")
    If respuesta = vbNo Then Exit Sub
    
    
    archivo = Application.GetOpenFilename("Archivos XML (*.xml), *.xml", , "Seleccionar archivo XML")
    If archivo = "False" Then Exit Sub
    If LCase(Trim(Right(archivo, 4))) <> ".xml" Then
        MsgBox "El archivo seleccionado no es un archivo XML.", vbExclamation
        Exit Sub
    End If
    
    
    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    xmlDoc.async = False
    xmlDoc.Load archivo
    If xmlDoc.parseError.ErrorCode <> 0 Then
        MsgBox "Error al cargar el archivo XML: " & xmlDoc.parseError.Reason, vbExclamation
        Exit Sub
    End If
    
    
    Set resultNodes = xmlDoc.SelectNodes("/report/report/results/result")
    If resultNodes.Length = 0 Then
        MsgBox "No se encontraron registros en el XML.", vbExclamation
        Exit Sub
    End If
    
    
    Set dict = CreateObject("Scripting.Dictionary")
    For Each header In tbl.HeaderRowRange.Cells
        dict(Trim(header.value)) = header.Column - tbl.Range.Cells(1, 1).Column + 1
    Next header
    
    
    requiredFields = Array("Severidad", "Nombre de vulnerabilidad", "Salidas de herramienta", "IPv4 Interna", "Puerto")
    For Each field In requiredFields
        If Not dict.exists(field) Then
            MsgBox "La columna "
            Exit Sub
        End If
    Next field
    
 
With regex
    .pattern = "\D"
    .Global = True
End With


For Each resultNode In resultNodes
    Set fila = tbl.ListRows.Add
    On Error Resume Next
    fila.Range.Cells(1, dict("Severidad")).value = Trim(resultNode.SelectSingleNode("severity").text)
    fila.Range.Cells(1, dict("Nombre de vulnerabilidad")).value = Trim(resultNode.SelectSingleNode("name").text)
    fila.Range.Cells(1, dict("Salidas de herramienta")).value = Trim(resultNode.SelectSingleNode("description").text)
    fila.Range.Cells(1, dict("IPv4 Interna")).value = Trim(resultNode.SelectSingleNode("host").text)

    
    fila.Range.Cells(1, dict("Puerto")).value = regex.Replace(Trim(resultNode.SelectSingleNode("port").text), "")

    On Error GoTo 0
Next resultNode

    MsgBox "Datos cargados con ?xito.", vbInformation
End Sub

' Helper function to get node text safely
Private Function GetNodeTextFromNode(ByVal parentNode As Object, ByVal xPathQuery As String) As String
    Dim selectedNode As Object
    On Error Resume Next ' Temporarily ignore errors if node not found
    Set selectedNode = parentNode.SelectSingleNode(xPathQuery)
    If Err.Number <> 0 Or selectedNode Is Nothing Then
        GetNodeTextFromNode = "N/A"
        Err.Clear
    Else
        GetNodeTextFromNode = Trim(selectedNode.text)
    End If
    On Error GoTo 0 ' Restore default error handling
    Set selectedNode = Nothing
End Function



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
    Dim fila As ListRow
    Dim filaExistente As ListRow
    Dim lrExisting As ListRow
    Dim colCSVIndex As Integer
    
    Dim j As Long
    Dim celdaEnTabla As Boolean
    Dim t As ListObject
    Dim mensaje As String
    Dim respuesta As VbMsgBoxResult

    
    Dim csvTarget As String
    Dim csvAffects As String
    Dim csvName As String
    Dim identificadorDeteccion As String
    Dim affectsProcessed As String

    
    Dim registroEncontrado As Boolean
    Dim fechaActual As String
    Dim colIdxIdDeteccion As Long, colIdxTipoOrigen As Long, colIdxIdVuln As Long, colIdxNomVuln As Long
    Dim colIdxConteo As Long, colIdxFecha As Long
    Dim match1 As Boolean, match2 As Boolean, match3 As Boolean, match4 As Boolean
    Dim conteoActual As Variant
    Dim nuevoConteo As Long

    
    Dim columnasFaltantes As String
    Dim colNombre As String

    
    On Error GoTo ErrorHandler

    
    fechaActual = Format(Date, "dd/mm/yyyy")

    
    Set wb = ThisWorkbook
    Set ws = wb.ActiveSheet

    
    celdaEnTabla = False
    For Each t In ws.ListObjects
        If Not t.DataBodyRange Is Nothing Then
            
            On Error Resume Next
            Dim tblRangeAddress As String
            tblRangeAddress = t.Range.Address
            If Err.Number = 0 Then
                 If Not Intersect(ActiveCell, t.Range) Is Nothing Then
                    Set tbl = t
                    celdaEnTabla = True
                    Exit For
                 End If
            Else
                Debug.Print "Advertencia: Error al acceder al rango de la tabla "
                Err.Clear
            End If
            On Error GoTo ErrorHandler
        End If
    Next t


    
    If Not celdaEnTabla Then
        MsgBox "La celda seleccionada no est� dentro de una tabla v�lida.", vbExclamation
        Exit Sub
    End If

    
    columnasFaltantes = ""
    


    
    On Error GoTo ErrorHandler

    
    CheckColumnExists tbl, "Identificador de detecci�n usado", columnasFaltantes, colIdxIdDeteccion
    CheckColumnExists tbl, "Tipo de origen", columnasFaltantes, colIdxTipoOrigen
    CheckColumnExists tbl, "Identificador original de la vulnerabilidad", columnasFaltantes, colIdxIdVuln
    CheckColumnExists tbl, "Nombre original de la vulnerabilidad", columnasFaltantes, colIdxNomVuln
    
    CheckColumnExists tbl, "Conteo de detecci�n", columnasFaltantes, colIdxConteo
    CheckColumnExists tbl, "�ltima fecha de detecci�n", columnasFaltantes, colIdxFecha

    
If Len(columnasFaltantes) > 0 Then
    MsgBox "Error Cr�tico: La(s) siguiente(s) columna(s) requerida(s) no existe(n) en la tabla: " & _
           columnasFaltantes & vbCrLf & vbCrLf & _
           "Por favor, aseg�rese de que estas columnas existan con el nombre EXACTO (incluyendo tildes, may�sculas/min�sculas si aplica y sin espacios extra).", _
           vbCritical, "Columnas Faltantes Espec�ficas"
    Exit Sub
End If

On Error GoTo ErrorHandler

mensaje = "Esta macro cargar� datos de Acunetix en la hoja actual." & vbCrLf & _
          "Verificar� duplicados basados en las 4 columnas clave." & vbCrLf & vbCrLf & _
          "- Si es nuevo: Agrega registro, Conteo = 1, Fecha = Hoy." & vbCrLf & _
          "- Si existe: Incrementa el conteo." & vbCrLf & vbCrLf & _
          "�Desea continuar?"

respuesta = MsgBox(mensaje, vbYesNo + vbQuestion, "Confirmaci�n de Carga con Verificaci�n")
If respuesta = vbNo Then Exit Sub


    
    archivos = Application.GetOpenFilename("Archivos CSV (*.csv), *.csv", MultiSelect:=True, Title:="Seleccionar archivos CSV de Acunetix")
    If IsArray(archivos) = False Then Exit Sub

    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    
    For i = LBound(archivos) To UBound(archivos)
        
        
        On Error Resume Next
        Set wbCSV = Workbooks.Open(fileName:=archivos(i), Local:=True, ReadOnly:=True, Format:=6, Delimiter:=",")
        If Err.Number <> 0 Then
            MsgBox "Error al abrir el archivo CSV: "
            Err.Clear
            On Error GoTo ErrorHandler
            GoTo SiguienteArchivo
        End If
        On Error GoTo ErrorHandler

        Set wsCSV = wbCSV.Sheets(1)

        
        If wsCSV.FilterMode Then wsCSV.ShowAllData
        If Not wsCSV.UsedRange Is Nothing Then
             If wsCSV.UsedRange.Rows.Count > 0 Then
                 encabezados = wsCSV.UsedRange.Rows(1).value
             Else
                 MsgBox "El archivo CSV "
                 GoTo CerrarYSaltar
             End If
             If wsCSV.UsedRange.Rows.Count > 1 Then
                  
                  csvData = wsCSV.UsedRange.Offset(1, 0).Resize(wsCSV.UsedRange.Rows.Count - 1, wsCSV.UsedRange.Columns.Count).value
             Else
                  csvData = Null
             End If
        Else
             MsgBox "El archivo CSV "
             GoTo CerrarYSaltar
        End If

        
        wbCSV.Close False
        Set wsCSV = Nothing
        Set wbCSV = Nothing

        If IsNull(csvData) Then GoTo SiguienteArchivo

        
        For j = 1 To UBound(csvData, 1)
            
            csvTarget = ""
            csvAffects = ""
            csvName = ""
            identificadorDeteccion = ""
            affectsProcessed = ""

            
            Dim colTargetIdx As Integer: colTargetIdx = 0
            Dim colAffectsIdx As Integer: colAffectsIdx = 0
            Dim colNameIdx As Integer: colNameIdx = 0
            On Error Resume Next
            For colCSVIndex = 1 To UBound(encabezados, 2)
                Select Case Trim(CStr(encabezados(1, colCSVIndex)))
                    Case "Target": colTargetIdx = colCSVIndex
                    Case "Affects": colAffectsIdx = colCSVIndex
                    Case "Name": colNameIdx = colCSVIndex
                End Select
            Next colCSVIndex
            Err.Clear
            On Error GoTo ErrorHandler

            
            If colTargetIdx = 0 Or colAffectsIdx = 0 Or colNameIdx = 0 Then
                 Debug.Print "Advertencia: Faltan columnas ("
                 GoTo SiguienteFilaCSV
            End If

            
            On Error Resume Next
            csvTarget = CStr(csvData(j, colTargetIdx))
            csvAffects = CStr(csvData(j, colAffectsIdx))
            csvName = CStr(csvData(j, colNameIdx))
            If Err.Number <> 0 Then
                Debug.Print "Advertencia: Error leyendo datos de fila " & j & " en CSV: " & Mid(archivos(i), InStrRev(archivos(i), "\") + 1) & ". Usando valores vac�os."
                csvTarget = "": csvAffects = "": csvName = ""
                Err.Clear
            End If
            On Error GoTo ErrorHandler


            
            If Right(csvTarget, 1) = "/" Then
                If Left(csvAffects, 1) = "/" Then
                    If Len(csvAffects) > 1 Then affectsProcessed = Mid(csvAffects, 2) Else affectsProcessed = ""
                Else
                    affectsProcessed = csvAffects
                End If
            Else
                 
                 affectsProcessed = csvAffects
            End If
             
            identificadorDeteccion = csvTarget & affectsProcessed


            
            registroEncontrado = False
            Set filaExistente = Nothing
            If tbl.ListRows.Count > 0 Then
                For Each lrExisting In tbl.ListRows
                    match1 = False: match2 = False: match3 = False: match4 = False
                    On Error Resume Next
                    
                    match1 = (CStr(lrExisting.Range(1, colIdxIdDeteccion).value) = identificadorDeteccion)
                    match2 = (CStr(lrExisting.Range(1, colIdxTipoOrigen).value) = "Acunetix")
                    match3 = (CStr(lrExisting.Range(1, colIdxIdVuln).value) = csvName)
                    match4 = (CStr(lrExisting.Range(1, colIdxNomVuln).value) = csvName)
                    
                    If Err.Number <> 0 Then Err.Clear
                    On Error GoTo ErrorHandler

                    If match1 And match2 And match3 And match4 Then
                        registroEncontrado = True
                        Set filaExistente = lrExisting
                        Exit For
                    End If
                Next lrExisting
            End If

            
            If registroEncontrado Then
                
                
                On Error Resume Next
                conteoActual = filaExistente.Range(1, colIdxConteo).value
                If IsNumeric(conteoActual) And Not IsEmpty(conteoActual) And Not IsNull(conteoActual) Then
                    nuevoConteo = CLng(conteoActual) + 1
                Else
                    nuevoConteo = 1
                End If
                filaExistente.Range(1, colIdxConteo).value = nuevoConteo
                If Err.Number <> 0 Then
                    Debug.Print "Advertencia: No se pudo actualizar "
                    Err.Clear
                End If
                On Error GoTo ErrorHandler

                
                On Error Resume Next
                filaExistente.Range(1, colIdxFecha).value = CDate(fechaActual)
                 If Err.Number <> 0 Then
                    Err.Clear
                    filaExistente.Range(1, colIdxFecha).value = fechaActual
                    If Err.Number <> 0 Then
                       Debug.Print "Advertencia: No se pudo actualizar "
                       Err.Clear
                    End If
                 End If
                On Error GoTo ErrorHandler

            Else
                
                 On Error Resume Next
                 Set fila = tbl.ListRows.Add(AlwaysInsert:=True)
                 If Err.Number <> 0 Then
                    MsgBox "Error Cr�tico al intentar agregar una nueva fila a la tabla "
                    Err.Clear
                    GoTo CleanupAndExit
                 End If
                 On Error GoTo ErrorHandler

                
                On Error Resume Next
                fila.Range(1, colIdxIdDeteccion).value = identificadorDeteccion
                If Err.Number <> 0 Then Debug.Print "Err escribiendo IdDeteccion: " & Err.Description: Err.Clear
                fila.Range(1, colIdxTipoOrigen).value = "Acunetix"
                 If Err.Number <> 0 Then Debug.Print "Err escribiendo TipoOrigen: " & Err.Description: Err.Clear
                fila.Range(1, colIdxIdVuln).value = csvName
                 If Err.Number <> 0 Then Debug.Print "Err escribiendo IdVuln: " & Err.Description: Err.Clear
                fila.Range(1, colIdxNomVuln).value = csvName
                 If Err.Number <> 0 Then Debug.Print "Err escribiendo NomVuln: " & Err.Description: Err.Clear
                fila.Range(1, colIdxConteo).value = 1
                 If Err.Number <> 0 Then Debug.Print "Err escribiendo Conteo: " & Err.Description: Err.Clear
                fila.Range(1, colIdxFecha).value = CDate(fechaActual)
                If Err.Number <> 0 Then
                    Err.Clear
                    fila.Range(1, colIdxFecha).value = fechaActual
                     If Err.Number <> 0 Then Debug.Print "Err escribiendo Fecha: " & Err.Description: Err.Clear
                End If
                On Error GoTo ErrorHandler

                Set fila = Nothing
            End If

            Set filaExistente = Nothing

SiguienteFilaCSV:
        Next j

CerrarYSaltar:
        If Not wbCSV Is Nothing Then
             If wbCSV.Name = Mid(archivos(i), InStrRev(archivos(i), "\") + 1) Then
                wbCSV.Close False
             End If
        End If
        Set wsCSV = Nothing
        Set wbCSV = Nothing

SiguienteArchivo:
    Next i

CleanupAndExit:
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    
    If Err.Number = 0 Then
       MsgBox "Proceso completado. Se verificaron duplicados y se actualizaron/agregaron registros en "
    End If
    
    Set tbl = Nothing: Set ws = Nothing: Set wb = Nothing
    Set wbCSV = Nothing: Set wsCSV = Nothing
    Set fila = Nothing: Set filaExistente = Nothing: Set lrExisting = Nothing: Set t = Nothing
    Exit Sub

ErrorHandler:
    
    MsgBox "Ocurri� un error inesperado:" & vbCrLf & _
           "N�mero de Error: " & Err.Number & vbCrLf & _
           "Descripci�n: " & Err.Description & vbCrLf & _
           "Fuente: " & Err.Source & vbCrLf & _
           "Puede haber ocurrido en el archivo: " & IIf(i > 0 And i <= UBound(archivos), Mid(archivos(i), InStrRev(archivos(i), "\") + 1), "N/A") & _
           ", Fila CSV (aprox): " & j, _
           vbCritical, "Error en Macro"

    
    If Not wbCSV Is Nothing Then
        On Error Resume Next
        wbCSV.Close False
        On Error GoTo 0
    End If

    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    
    Set tbl = Nothing: Set ws = Nothing: Set wb = Nothing
    Set wbCSV = Nothing: Set wsCSV = Nothing
    Set fila = Nothing: Set filaExistente = Nothing: Set lrExisting = Nothing: Set t = Nothing
    

End Sub

    Sub CheckColumnExists(ByVal tblToCheck As ListObject, ByVal columnName As String, ByRef missingList As String, ByRef outIndex As Long)
        On Error Resume Next
        outIndex = 0
        outIndex = tblToCheck.ListColumns(columnName).Index
        If Err.Number <> 0 Then
            
            If Len(missingList) > 0 Then missingList = missingList & ", "
            missingList = missingList & ""
            Err.Clear
        End If
        On Error GoTo 0
    End Sub


' Helper function to strip HTML tags
Private Function StripHTML(ByVal htmlText As String) As String
    Dim regexHTML As Object
    Set regexHTML = CreateObject("VBScript.RegExp")
    With regexHTML
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .pattern = "<[^>]+>" ' Matches any HTML tag
    End With
    StripHTML = regexHTML.Replace(htmlText, "")
    StripHTML = Replace(StripHTML, "�", " ") ' Replace non-breaking space
    StripHTML = Replace(StripHTML, vbCrLf & vbCrLf, vbCrLf) ' Reduce multiple newlines
    StripHTML = Trim(StripHTML)
    Set regexHTML = Nothing
End Function


Sub CYB040_ResaltarFalsosPositivosEnVerde()
   Dim ws As Worksheet
    Dim tbl As ListObject
    Dim celda As Range
    Dim valoresTabla As Object
    Dim columnaIndex As Integer
    Dim encontrada As Boolean
    
    
    Set valoresTabla = CreateObject("Scripting.Dictionary")
    
    
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

    
    If Not encontrada Then
        MsgBox "No se encontr? la tabla "
        Exit Sub
    End If

    
    On Error Resume Next
    columnaIndex = tbl.ListColumns("Vulnerability Name").Index
    On Error GoTo 0
    If columnaIndex = 0 Then
        MsgBox "La columna "
        Exit Sub
    End If

    
    Dim celdaTabla As Range
    For Each celdaTabla In tbl.ListColumns(columnaIndex).DataBodyRange
        valoresTabla(celdaTabla.value) = True
    Next celdaTabla

    
    Dim coincidencias As Boolean
    coincidencias = False
    For Each celda In Selection
        If valoresTabla.exists(celda.value) Then
            celda.Interior.Color = RGB(0, 255, 0)
            coincidencias = True
        End If
    Next celda

    
    If coincidencias Then
        MsgBox "Se han resaltado las celdas seleccionadas que coinciden con valores en "
    Else
        MsgBox "No hay coincidencias en la tabla.", vbExclamation
    End If
End Sub


' Helper function to strip HTML tags and format text
Private Function FormatDetailsText(ByVal rawText As String) As String
    Dim tempText As String
    Dim regexHTML As Object
    
    Set regexHTML = CreateObject("VBScript.RegExp")
    With regexHTML
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .pattern = "<[^>]+>" ' Matches any HTML tag
    End With
    
    tempText = regexHTML.Replace(rawText, "") ' Remove HTML tags first
    Set regexHTML = Nothing
    
    ' Replace specific characters with themselves + newline
    


   Set regexColon = CreateObject("VBScript.RegExp")
    With regexColon
        .Global = True
        .MultiLine = True
        .IgnoreCase = True ' Though not strictly needed for this pattern
        ' Pattern: Match a colon (:) NOT followed by:
        '   - "//" (to protect http://, ftp:// etc.)
        '   - two digits (to protect times like 16:51 or 01:51)
        .pattern = ":(?!\/\/|\d{2})"
    End With
    tempText = regexColon.Replace(tempText, ":" & vbCrLf)
    Set regexColon = Nothing
    
    
    tempText = Replace(tempText, ":http", ": http" & vbCrLf)
    tempText = Replace(tempText, ";", ";" & vbCrLf)
    
    tempText = Replace(tempText, "  ", " " & vbCrLf)
    
    ' Clean up potential extra spaces around newlines or at the start/end
    tempText = Replace(tempText, vbCrLf & " ", vbCrLf) ' Remove space after newline
    tempText = Replace(tempText, " " & vbCrLf, vbCrLf) ' Remove space before newline
    tempText = Replace(tempText, vbCrLf & vbCrLf, vbCrLf) ' Consolidate multiple newlines
    
    ' Replace common HTML entities that might remain
    tempText = Replace(tempText, "�", " ")
    tempText = Replace(tempText, "<", "<")
    tempText = Replace(tempText, ">", ">")
    tempText = Replace(tempText, "&", "&")
    tempText = Replace(tempText, "'", "'")

    FormatDetailsText = Trim(tempText)
End Function


Sub CYB036_CargarResultados_DatosDesdeXMLAcunetix_v5()
    Dim wb As Workbook, ws As Worksheet, tbl As ListObject
    Dim archivo As String
    Dim xmlDoc As Object, reportItemNodes As Object, reportItemNode As Object
    Dim respuesta As Integer, mensaje As String
    Dim fila As ListRow
    Dim dict As Object, header As Range
    Dim requiredFields As Variant, field As Variant
    
    Dim regexLi As Object, liMatches As Object, liMatch As Object
    Dim regexUrl As Object, urlMatch As Object
    
    Dim detailsNode As Object
    Dim detailsRawText As String ' Raw text from XML node
    Dim detailsFormattedText As String ' Text after stripping HTML and formatting
    Dim firstUrlFound As String
    Dim cweStr As String
    Dim cweNodes As Object, cweNode As Object
    Dim firstItem As Boolean

    Set wb = ThisWorkbook
    Set ws = wb.ActiveSheet

    ' --- Check if active cell is within a table ---
    Dim celdaEnTabla As Boolean, t As ListObject
    celdaEnTabla = False
    If ws.ListObjects.Count = 0 Then
        MsgBox "No hay tablas en la hoja activa.", vbExclamation
        Exit Sub
    End If
    For Each t In ws.ListObjects
        If Not t.DataBodyRange Is Nothing Then
            If Not Intersect(ActiveCell, t.Range) Is Nothing Then
                Set tbl = t
                celdaEnTabla = True
                Exit For
            End If
        Else
            If Not Intersect(ActiveCell, t.HeaderRowRange) Is Nothing Then
                Set tbl = t
                celdaEnTabla = True
                Exit For
            End If
        End If
    Next t

    If Not celdaEnTabla Then
        MsgBox "La celda seleccionada no est� dentro de una tabla o la tabla est� vac�a." & vbCrLf & _
               "Por favor, seleccione una celda dentro de la tabla de destino.", vbExclamation
        Exit Sub
    End If

    ' --- Confirmation ---
    mensaje = "�Est� seguro que desea cargar datos del archivo XML de Acunetix en la tabla '" & tbl.Name & "'?"
    respuesta = MsgBox(mensaje, vbYesNo + vbQuestion, "Confirmaci�n")
    If respuesta = vbNo Then Exit Sub

    ' --- Get XML File ---
    archivo = Application.GetOpenFilename("Archivos XML (*.xml), *.xml", , "Seleccionar archivo XML de Acunetix")
    If archivo = "False" Then Exit Sub
    If LCase(Trim(Right(archivo, 4))) <> ".xml" Then
        MsgBox "El archivo seleccionado no es un archivo XML.", vbExclamation
        Exit Sub
    End If

    ' --- Load XML Document ---
    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    xmlDoc.async = False
    xmlDoc.SetProperty "SelectionLanguage", "XPath"
    xmlDoc.Load archivo
    If xmlDoc.parseError.ErrorCode <> 0 Then
        MsgBox "Error al cargar el archivo XML: " & xmlDoc.parseError.Reason, vbExclamation
        Exit Sub
    End If

    ' --- Get ReportItem Nodes ---
    Set reportItemNodes = xmlDoc.SelectNodes("/ScanGroup/Scan/ReportItems/ReportItem")
    If reportItemNodes Is Nothing Or reportItemNodes.Length = 0 Then
        MsgBox "No se encontraron registros de vulnerabilidades (ReportItem) en el XML. XPath: /ScanGroup/Scan/ReportItems/ReportItem", vbExclamation
        Exit Sub
    End If

    ' --- Create Dictionary of Column Headers ---
    Set dict = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    For Each header In tbl.HeaderRowRange.Cells
        If Trim(header.value) <> "" Then
            dict(Trim(header.value)) = header.Column - tbl.Range.Cells(1, 1).Column + 1
        End If
    Next header
    On Error GoTo 0
    If dict.Count = 0 Then
        MsgBox "La tabla seleccionada no tiene cabeceras v�lidas.", vbExclamation
        Exit Sub
    End If

    ' --- Check for Required Fields in the Table ---
    requiredFields = Array("Explicaci�n t�cnica", "Nombre original de la vulnerabilidad", "Identificador de detecci�n usado", "Tipo de origen", "Identificador original de la vulnerabilidad", "CWE", "Salidas de herramienta")
    For Each field In requiredFields
        If Not dict.exists(field) Then
            MsgBox "La columna '" & field & "' no se encuentra en la tabla '" & tbl.Name & "'. Aseg�rese de que todas las columnas requeridas existan.", vbCritical
            Exit Sub
        End If
    Next field

    ' --- Initialize Regex for <li> tags ---
    Set regexLi = CreateObject("VBScript.RegExp")
    With regexLi
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .pattern = "<li>(.*?)</li>"
    End With
    
    ' --- Initialize Regex for URLs ---
    Set regexUrl = CreateObject("VBScript.RegExp")
    With regexUrl
        .Global = False
        .IgnoreCase = True
        .pattern = "(https?://[^\s<>"";]+)" ' Added semicolon to excluded chars for URL
    End With


    Application.ScreenUpdating = False

    ' --- Process Each ReportItem Node ---
    For Each reportItemNode In reportItemNodes
        Set fila = tbl.ListRows.Add
        
        ' Explicaci�n t�cnica -> <Description>
        fila.Range.Cells(1, dict("Explicaci�n t�cnica")).value = GetNodeTextFromNode(reportItemNode, "Description")

        ' Nombre original de la vulnerabilidad -> <ModuleName>
        fila.Range.Cells(1, dict("Nombre original de la vulnerabilidad")).value = GetNodeTextFromNode(reportItemNode, "ModuleName")

        ' Salidas de herramienta -> Formatted plain text from <Details>
        detailsFormattedText = "N/A"
        Set detailsNode = Nothing
        On Error Resume Next
        Set detailsNode = reportItemNode.SelectSingleNode("Details")
        On Error GoTo 0
        If Not detailsNode Is Nothing Then
            detailsRawText = Trim(detailsNode.text) ' Get raw text from CDATA
            If Len(detailsRawText) = 0 And Len(Trim(detailsNode.XML)) > 0 Then ' Fallback to full XML if .text is empty
                detailsRawText = Trim(detailsNode.XML)
            End If
            
            If Len(detailsRawText) > 0 Then
                detailsFormattedText = FormatDetailsText(detailsRawText)
            End If
        End If
        If Len(detailsFormattedText) = 0 Or detailsFormattedText = "N/A" Then
             detailsFormattedText = "N/A (no <Details> content)"
        End If
        fila.Range.Cells(1, dict("Salidas de herramienta")).value = detailsFormattedText


        ' Identificador de detecci�n usado -> First URL from <li> within <Details>
        firstUrlFound = "N/A"
        Set detailsNode = Nothing
        On Error Resume Next
        Set detailsNode = reportItemNode.SelectSingleNode("Details")
        On Error GoTo 0

        If Not detailsNode Is Nothing Then
            detailsRawText = Trim(detailsNode.text)
            If Len(detailsRawText) > 0 Then
                Set liMatches = regexLi.Execute(detailsRawText) ' Execute on raw text to find <li>
                If liMatches.Count > 0 Then
                    For Each liMatch In liMatches
                        Dim liContent As String
                        liContent = Trim(liMatch.SubMatches(0)) ' Content of <li>
                        
                        Set urlMatch = regexUrl.Execute(liContent) ' Find URL within this <li>
                        If urlMatch.Count > 0 Then
                            firstUrlFound = Trim(urlMatch(0).value)
                            Exit For
                        End If
                    Next liMatch
                End If
            End If
        End If
        fila.Range.Cells(1, dict("Identificador de detecci�n usado")).value = firstUrlFound


        ' Tipo de origen -> "Acunetix"
        fila.Range.Cells(1, dict("Tipo de origen")).value = "Acunetix"

        ' Identificador original de la vulnerabilidad -> <Name>
        fila.Range.Cells(1, dict("Identificador original de la vulnerabilidad")).value = GetNodeTextFromNode(reportItemNode, "Name")

        ' CWE -> <CWEList/CWE>
        cweStr = ""
        Set cweNodes = Nothing
        On Error Resume Next
        Set cweNodes = reportItemNode.SelectNodes("CWEList/CWE")
        On Error GoTo 0

        If Not cweNodes Is Nothing Then
            If cweNodes.Length > 0 Then
                firstItem = True
                For Each cweNode In cweNodes
                    Dim cweIdAttr As Object, cweVal As String
                    cweVal = ""
                    Set cweIdAttr = Nothing
                    On Error Resume Next
                    Set cweIdAttr = cweNode.Attributes.getNamedItem("id")
                    On Error GoTo 0

                    If Not cweIdAttr Is Nothing And Len(Trim(cweIdAttr.text)) > 0 Then
                        cweVal = "CWE-" & Trim(cweIdAttr.text)
                    ElseIf Len(Trim(cweNode.text)) > 0 Then
                        cweVal = Trim(cweNode.text)
                        If Not (UCase(Left(cweVal, 4)) = "CWE-") And IsNumeric(Replace(cweVal, "CWE-", "")) Then
                            If IsNumeric(cweVal) Then cweVal = "CWE-" & cweVal
                        End If
                    End If

                    If Len(cweVal) > 0 Then
                        If Not firstItem Then
                            cweStr = cweStr & "; "
                        End If
                        cweStr = cweStr & cweVal
                        firstItem = False
                    End If
                Next cweNode
            End If
        End If
        If Len(cweStr) = 0 Then cweStr = "N/A"
        fila.Range.Cells(1, dict("CWE")).value = cweStr

        ' Clear objects for next iteration
        Set detailsNode = Nothing
        Set cweNodes = Nothing
        Set cweNode = Nothing
        Set liMatches = Nothing
        Set urlMatch = Nothing
    Next reportItemNode

    Application.ScreenUpdating = True
    MsgBox "Datos cargados con �xito desde el archivo Acunetix (v5).", vbInformation

    ' Clean up
    Set xmlDoc = Nothing
    Set reportItemNodes = Nothing
    Set reportItemNode = Nothing
    Set dict = Nothing
    Set tbl = Nothing
    Set ws = Nothing
    Set wb = Nothing
    Set regexLi = Nothing
    Set regexUrl = Nothing
End Sub

Sub CYB041_IrACatalogoVulnerabilidad()
Attribute CYB041_IrACatalogoVulnerabilidad.VB_ProcData.VB_Invoke_Func = "G\n14"

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

    
    Set wsOrigen = ActiveSheet
    Set wsCatalogo = ThisWorkbook.Sheets("Catalogo vulnerabilidades")
    
    
    If wsOrigen.ListObjects.Count = 0 Then
        MsgBox "No se encontr� una tabla en la hoja actual.", vbExclamation, "Error"
        Exit Sub
    End If
    Set tblOrigen = wsOrigen.ListObjects(1)
    
    
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
    
    tipoOrigen = tblOrigen.ListColumns("Tipo de origen").DataBodyRange.Cells(rngCeldaActual.row - tblOrigen.DataBodyRange.row + 1, 1).value
    idVulnerabilidad = tblOrigen.ListColumns("Identificador original de la vulnerabilidad").DataBodyRange.Cells(rngCeldaActual.row - tblOrigen.DataBodyRange.row + 1, 1).value
    
    If tipoOrigen = "" Or IsEmpty(idVulnerabilidad) Then
        MsgBox "Falta el Tipo de Origen o Identificador en la fila seleccionada.", vbExclamation, "Error"
        Exit Sub
    End If
    
    If Not dictColumnas.exists(tipoOrigen) Then
        MsgBox "El tipo de origen "
        Exit Sub
    End If
    
    colBusqueda = dictColumnas(tipoOrigen)
    
    On Error Resume Next
    Set tblCatalogo = wsCatalogo.ListObjects("Tbl_Catalogo_vulnerabilidades")
    On Error GoTo 0
    
    If tblCatalogo Is Nothing Then
        MsgBox "No se encontr� la tabla "
        Exit Sub
    End If
    
    Set rngBusqueda = tblCatalogo.ListColumns(colBusqueda).DataBodyRange
    Set celdaEncontrada = rngBusqueda.Find(What:=idVulnerabilidad, LookAt:=xlWhole)
    
    If Not celdaEncontrada Is Nothing Then
        wsCatalogo.Activate
        celdaEncontrada.EntireRow.Select
        MsgBox "Registro encontrado. Se ha seleccionado la fila correspondiente en el cat�logo.", vbInformation, "�xito"
    Else
        respuesta = MsgBox("No se encontr� el identificador en el cat�logo. �Deseas agregarlo?", vbYesNo + vbQuestion, "Agregar nuevo registro")
        If respuesta = vbYes Then
            Set nuevaFila = tblCatalogo.ListRows.Add
            
            
            On Error Resume Next
            tblCatalogo.ListColumns(colBusqueda).DataBodyRange.Cells(nuevaFila.Index, 1).value = idVulnerabilidad
            On Error GoTo 0
            
            
            fechaActual = Format(Now, "dd/mm/yyyy")
            
            On Error Resume Next
            tblCatalogo.ListColumns("SourceDetection").DataBodyRange.Cells(nuevaFila.Index, 1).value = tipoOrigen
            tblCatalogo.ListColumns("LastEditedBy").DataBodyRange.Cells(nuevaFila.Index, 1).value = "Default System"
            tblCatalogo.ListColumns("LastUpdateDate").DataBodyRange.Cells(nuevaFila.Index, 1).value = fechaActual
            On Error GoTo 0
            
            
            Application.GoTo Reference:=tblCatalogo.ListRows(nuevaFila.Index).Range, Scroll:=True
            
            MsgBox "Nuevo registro agregado al cat�logo. Se ha seleccionado la fila correspondiente.", vbInformation, "�xito"
        End If
    End If
End Sub

Sub CYB042_MarcarMultiplesEnCatalogoVulnerabilidad()

    Dim wsOrigen As Worksheet, wsCatalogo As Worksheet
    Dim tblOrigen As ListObject, tblCatalogo As ListObject
    Dim rngCeldaActual As Range, rngSeleccion As Range
    Dim idVulnerabilidad As Variant, tipoOrigen As String
    Dim colBusqueda As String, rngBusqueda As Range, celdaEncontrada As Range
    Dim dictColumnas As Object
    Dim respuesta As VbMsgBoxResult
    Dim colSourceDetection As Long, colLastEditedBy As Long, colLastUpdateDate As Long
    Dim fechaActual As String
    Dim fila As Range

    Set wsOrigen = ActiveSheet
    Set wsCatalogo = ThisWorkbook.Sheets("Catalogo vulnerabilidades")
    
    ' Verificar que haya una tabla en la hoja activa
    If wsOrigen.ListObjects.Count = 0 Then
        MsgBox "No se encontr� una tabla en la hoja actual.", vbExclamation, "Error"
        Exit Sub
    End If
    Set tblOrigen = wsOrigen.ListObjects(1)
    
    ' Configurar el diccionario para mapear los tipos de origen a las columnas
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
    
    ' Comprobar que la celda seleccionada est� dentro de la tabla
    Set rngSeleccion = Selection
    If Intersect(rngSeleccion, tblOrigen.DataBodyRange) Is Nothing Then
        MsgBox "Selecciona celdas dentro de la tabla de vulnerabilidades.", vbExclamation, "Error"
        Exit Sub
    End If

    ' Iterar sobre todas las celdas seleccionadas
    For Each rngCeldaActual In rngSeleccion
        If Not Intersect(rngCeldaActual, tblOrigen.DataBodyRange) Is Nothing Then
            tipoOrigen = tblOrigen.ListColumns("Tipo de origen").DataBodyRange.Cells(rngCeldaActual.row - tblOrigen.DataBodyRange.row + 1, 1).value
            idVulnerabilidad = tblOrigen.ListColumns("Identificador original de la vulnerabilidad").DataBodyRange.Cells(rngCeldaActual.row - tblOrigen.DataBodyRange.row + 1, 1).value
            
            If tipoOrigen = "" Or IsEmpty(idVulnerabilidad) Then
                MsgBox "Falta el Tipo de Origen o Identificador en la fila seleccionada.", vbExclamation, "Error"
                Exit Sub
            End If
            
            If Not dictColumnas.exists(tipoOrigen) Then
                MsgBox "El tipo de origen no es v�lido.", vbExclamation, "Error"
                Exit Sub
            End If
            
            colBusqueda = dictColumnas(tipoOrigen)
            
            ' Buscar el identificador de vulnerabilidad en el cat�logo
            On Error Resume Next
            Set tblCatalogo = wsCatalogo.ListObjects("Tbl_Catalogo_vulnerabilidades")
            On Error GoTo 0
            
            If tblCatalogo Is Nothing Then
                MsgBox "No se encontr� la tabla en el cat�logo.", vbExclamation, "Error"
                Exit Sub
            End If
            
            Set rngBusqueda = tblCatalogo.ListColumns(colBusqueda).DataBodyRange
            Set celdaEncontrada = rngBusqueda.Find(What:=idVulnerabilidad, LookAt:=xlWhole)
            
            ' Si se encuentra la vulnerabilidad en el cat�logo, marcar la celda correspondiente
            If Not celdaEncontrada Is Nothing Then
                celdaEncontrada.EntireRow.Cells(1).Interior.Color = RGB(255, 255, 0) ' Color amarillo
            End If
        End If
    Next rngCeldaActual

    MsgBox "Se han marcado las celdas correspondientes en el cat�logo.", vbInformation, "�xito"
End Sub



Sub CYB042_Estandarizar()
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim rng As Range, cell As Range
    Dim dict As Object
    Dim colIndex As Object
    Dim key As String
    Dim i As Long, j As Long
    
    
    Set ws = ActiveSheet
    Set dict = CreateObject("Scripting.Dictionary")
    Set colIndex = CreateObject("Scripting.Dictionary")
    
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    
    Dim stdCol As Integer
    stdCol = 0
    
    For i = 1 To lastCol
        If ws.Cells(1, i).value = "StandardVulnerabilityName" Then
            stdCol = i
        End If
        
        If Not IsEmpty(ws.Cells(1, i).value) Then
            colIndex(ws.Cells(1, i).value) = i
        End If
    Next i
    
    If stdCol = 0 Then
        MsgBox "La columna "
        Exit Sub
    End If
    
    
    For i = 2 To lastRow
        key = ws.Cells(i, stdCol).value
        If key <> "" Then
            If Not dict.exists(key) Then
                dict.Add key, CreateObject("Scripting.Dictionary")
            End If
            
            
            For Each colName In colIndex.Keys
                Dim colNum As Integer
                colNum = colIndex(colName)
                If ws.Cells(i, colNum).value <> "" Then
                    dict(key)(colName) = ws.Cells(i, colNum).value
                End If
            Next
        End If
    Next i
    
    
    For i = 2 To lastRow
        key = ws.Cells(i, stdCol).value
        If key <> "" And dict.exists(key) Then
            For Each colName In colIndex.Keys
                colNum = colIndex(colName)
                If ws.Cells(i, colNum).value = "" And dict(key).exists(colName) Then
                    ws.Cells(i, colNum).value = dict(key)(colName)
                End If
            Next
        End If
    Next i
    
    MsgBox "Estandarizaci?n completada.", vbInformation
End Sub




Sub CYB043_AplicarFormatoCondicional()
    Dim selectedRange As Range

    
    On Error Resume Next
    Set selectedRange = Selection.SpecialCells(xlCellTypeConstants)
    On Error GoTo 0

    
    If selectedRange Is Nothing Then
        MsgBox "No hay celdas seleccionadas."
        Exit Sub
    End If

    
    With selectedRange
        .FormatConditions.Add Type:=xlTextString, String:="CR�TICA", TextOperator:=xlContains
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Font
            .Color = RGB(255, 255, 255)
        End With
        With .FormatConditions(1).Interior
            .Color = RGB(112, 48, 160)
        End With

        .FormatConditions.Add Type:=xlTextString, String:="ALTA", TextOperator:=xlContains
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Font
            .Color = RGB(255, 255, 255)
        End With
        With .FormatConditions(1).Interior
            .Color = RGB(255, 0, 0)
        End With

        .FormatConditions.Add Type:=xlTextString, String:="MEDIA", TextOperator:=xlContains
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Font
            .Color = RGB(0, 0, 0)
        End With
        With .FormatConditions(1).Interior
            .Color = RGB(255, 255, 0)
        End With

        .FormatConditions.Add Type:=xlTextString, String:="BAJA", TextOperator:=xlContains
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Font
            .Color = RGB(255, 255, 255)
        End With
        With .FormatConditions(1).Interior
            .Color = RGB(0, 176, 80)
        End With

        .FormatConditions.Add Type:=xlTextString, String:="INFORMATIVA", TextOperator:=xlContains
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Font
            .Color = RGB(0, 0, 0)
        End With
        With .FormatConditions(1).Interior
            .Color = RGB(231, 230, 230)
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
    
    
    If Selection.Cells.Count = 0 Then
        MsgBox "Seleccione al menos una celda con una vulnerabilidad antes de ejecutar la macro.", vbExclamation, "Error"
        Exit Sub
    End If
    
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    
    For Each cell In Selection
        
        If Not IsEmpty(cell.value) Then
            Vulnerabilidad = cell.value
            
            
            Dim prompt As String
            prompt = ConstruirPrompt(Vulnerabilidad)
            
            
            JSONBody = "{""model"": ""llama3.2:1b"", ""prompt"": """ & Replace(prompt, """", "\""") & """, ""stream"": false}"
            
            
            With http
                .Open "POST", "http://localhost:11434/api/generate", False
                .setRequestHeader "Content-Type", "application/json"
                .Send JSONBody
                response = .responseText
            End With
            
            
            extractedResponse = ExtraerRespuesta(response)
            
            
            cell.value = extractedResponse
        End If
    Next cell
    
    
    Set http = Nothing
End Sub


Sub CYB060_LLLM_deepseek_r1_1_5b()
    Dim http As Object
    Dim JSONBody As String
    Dim response As String
    Dim Vulnerabilidad As String
    Dim extractedResponse As String
    Dim cell As Range
    
    
    If Selection.Cells.Count = 0 Then
        MsgBox "Seleccione al menos una celda con una vulnerabilidad antes de ejecutar la macro.", vbExclamation, "Error"
        Exit Sub
    End If
    
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    
    For Each cell In Selection
        
        If Not IsEmpty(cell.value) Then
            Vulnerabilidad = cell.value
            
            
            Dim prompt As String
            prompt = ConstruirPrompt(Vulnerabilidad)
            
            
            JSONBody = "{""model"": ""deepseek-r1:1.5b"", ""prompt"": """ & Replace(prompt, """", "\""") & """, ""stream"": false}"
            
            
            With http
                .Open "POST", "http://localhost:11434/api/generate", False
                .setRequestHeader "Content-Type", "application/json"
                .Send JSONBody
                response = .responseText
            End With
            
            
            extractedResponse = ExtraerRespuesta(response)
            
            
            cell.value = extractedResponse
        End If
    Next cell
    
    
    Set http = Nothing
End Sub





Function ConstruirPrompt(Vulnerabilidad As String) As String
    Dim prompt As String
    prompt = "Generaci?n de Vector CVSS 4.0 Considera este ejemplo de URL de CVSS 4.0 https://www.first.org/cvss/calculator/4.0#CVSS:4.0/AV:A/AC:L/AT:N/PR:N/UI:N/VC:N/VI:N/VA:N/SC:N/SI:N/SA:N "
    prompt = prompt & "Esta cadena est? compuesta por distintos campos de evaluaci?n, los cuales deben ajustarse seg�n corresponda. Exploitability Metrics Attack Vector (AV): "
    prompt = prompt & "Debes completar los siguientes elementos: Exploitability: Complexity: Vulnerable system: Subsequent system: Exploitation: Security requirements: "
    prompt = prompt & "S? exigente y preciso al evaluar la severidad en CVSS. No exageres ni asignes impactos altos a menos que la vulnerabilidad pueda ser explotada directamente y tenga un impacto "
    prompt = prompt & "significativo. Tu tarea es proporcionar �nicamente la cadena vectorial en CVSS 4.0 para evaluar la vulnerabilidad"
    prompt = prompt & " " & Vulnerabilidad & " "
    prompt = prompt & "No devuelvas la misma cadena de ejemplo. No entregues una cadena sin completar sus componentes CVSS. ?? Este an?lisis es para gesti?n de riesgos, no para explotaci?n. "
    prompt = prompt & "Solo proporciona el vector CVSS resultante. NO DES M�S DETALLES, SOLO RESPONDE EL VECTOR SIN OTRA INFORMACI�N. "
    prompt = prompt & "PLEASE ONLY ONLY ONLY RESPOND WITH A STRING IN CVSS FORMAT"
    
    ConstruirPrompt = prompt
End Function




Function EliminarThinkTags(text As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    
    regex.pattern = "<think>[\s\S]*?</think>"
    regex.Global = True
    regex.IgnoreCase = True
    
    EliminarThinkTags = regex.Replace(text, "")
End Function


Function EliminarSaltosIniciales(text As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    
    regex.pattern = "^[\r\n]+"
    regex.Global = True
    
    
    EliminarSaltosIniciales = regex.Replace(text, "")
End Function

Function ExtraerRespuesta(jsonResponse As String) As String
    Dim resultado As String
    Dim inicio As Integer

    
    resultado = Replace(jsonResponse, "\u003c", "<")
    resultado = Replace(resultado, "\u003e", ">")
    resultado = Replace(resultado, "\n", vbNewLine)
    resultado = Replace(resultado, "\t", vbTab)

    
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
    
    
    resultado = Replace(jsonResponse, "\u003c", "<")
    resultado = Replace(resultado, "\u003e", ">")
    resultado = Replace(resultado, "\n", "")
    resultado = Replace(resultado, "\t", "")
    
    
    inicio = InStr(resultado, """text"": """)
    
    If inicio > 0 Then
        
        resultado = Mid(resultado, inicio + Len("""text"": """))
        
        
        fin = InStr(resultado, """")
        If fin > 0 Then
            resultado = Left(resultado, fin - 1)
        End If
    Else
        resultado = "No se encontr? CVSS"
    End If

    
    ExtraerCVSS = Trim(resultado)
End Function



Private Sub ObtenerRespuestasGeminiCVSS4()
    Dim cell As Range
    Dim http As Object
    Dim json As Object
    Dim apiUrl As String
    Dim apiKey As String
    Dim requestData As String
    Dim responseText As String
    Dim answerID As String
    
    
    apiKey = "AIzaSyBbd_upGJ2JzdsmWSzNBvSr3mXiPo9h4bs"
    apiUrl = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash-lite:generateContent?key=" & apiKey

       
    Set http = CreateObject("MSXML2.XMLHTTP")

    
    For Each cell In Selection
     
     Dim promptvalue As String
     promptvalue = ConstruirPrompt(cell.value)
    
    
        
        requestData = "{""contents"": [{""parts"": [{""text"": """ & promptvalue & """}]}]}"


        
        With http
            .Open "POST", apiUrl, False
            .setRequestHeader "Content-Type", "application/json"
            .Send requestData
        End With

        
        If http.Status = 200 Then
            responseText = http.responseText
            Debug.Print "Response: " & responseText

            
            On Error Resume Next
            Set json = JsonConverter.ParseJson(responseText)
            On Error GoTo 0

            
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
            Debug.Print "Response: " & responseText
        End If

        
        cell.Offset(0, 1).value = ExtraerCVSS(responseText)
    Next cell

    
    Set http = Nothing
    Set json = Nothing

    MsgBox "Procesamiento completado.", vbInformation
End Sub





Sub CYB068_PrepararPromptDesdeSeleccion_DescripcionVuln_General()
    Dim celda As Range
    Dim listaVulnerabilidades As String
    Dim prompt As String

    
    listaVulnerabilidades = ""

    
    If TypeName(Selection) <> "Range" Then
        MsgBox "Por favor, selecciona un rango de celdas.", vbExclamation, "Selecci�n Inv�lida"
        Exit Sub
    End If

    
    For Each celda In Selection
        If Not IsEmpty(celda.value) Then
            If listaVulnerabilidades <> "" Then
                listaVulnerabilidades = listaVulnerabilidades & "," & vbCrLf
            End If
            listaVulnerabilidades = listaVulnerabilidades & Trim(CStr(celda.value))
        End If
    Next celda

    
    If listaVulnerabilidades <> "" Then
        prompt = "Redacta un p�rrafo t�cnico breve y conciso que describa la vulnerabilidad detectada, comenzando con la frase: -El sistema...-. Explica en qu� consiste la debilidad de seguridad de manera t�cnica. No incluyas escenarios de explotaci�n, ya que eso corresponde a otro campo. No describas c�mo se explota, solo en qu� consiste el problema. No menciones el nombre exacto de la vulnerabilidad; utiliza expresiones similares. Vulnerabilidades a describir en formato tabla: "
        prompt = prompt & " " & vbCrLf & Chr(10) & listaVulnerabilidades & vbCrLf & Chr(10)
        
        prompt = prompt & "Contexto de an�lisis: An�lisis de vulnerabilidades realizado en el entorno evaluado. Vulnerabilidad detectada mediante escaneo e interacciones en el sistema bajo revisi�n."

        
        CopiarAlPortapapeles prompt

        
        MsgBox "El prompt para descripci�n de vulnerabilidad ha sido copiado al portapapeles.", vbInformation, "Prompt Generado"
    Else
        MsgBox "No se encontraron valores en la selecci�n.", vbExclamation, "Sin Datos"
    End If
End Sub


Sub CYB069_PrepararPromptDesdeSeleccion_AmenazaVuln_General()
    Dim celda As Range
    Dim listaVulnerabilidades As String
    Dim prompt As String

    
    listaVulnerabilidades = ""

    
    If TypeName(Selection) <> "Range" Then
        MsgBox "Por favor, selecciona un rango de celdas.", vbExclamation, "Selecci�n Inv�lida"
        Exit Sub
    End If

    
    For Each celda In Selection
        If Not IsEmpty(celda.value) Then
            If listaVulnerabilidades <> "" Then
                listaVulnerabilidades = listaVulnerabilidades & "," & vbCrLf
            End If
            listaVulnerabilidades = listaVulnerabilidades & Trim(CStr(celda.value))
        End If
    Next celda

    
    If listaVulnerabilidades <> "" Then
        prompt = "Hola, por favor considera el siguiente ejemplo: Un atacante podr�a (obtener, realizar, ejecutar, visualizar, identificar, listar)... Algunos posibles vectores de ataque adicionales asociados con esta amenaza incluyen (pueden ser algunos de los siguientes, pero no necesariamente todos): � Malware: Un malware dise�ado para automatizar intentos de fuerza bruta podr�a explotar la vulnerabilidad para... � Usuario malintencionado: Un usuario dentro del entorno con conocimiento de la vulnerabilidad podr�a aprovecharla para... � Personal interno: Un empleado con acceso y conocimientos t�cnicos podr�a, intencionalmente o por error,... � Delincuente cibern�tico: Un atacante externo en busca de vulnerabilidades podr�a intentar explotar esta debilidad para... Instrucciones adicionales: "
        prompt = prompt & "1. Pregunta si el sistema es interno o accesible externamente para determinar los vectores de ataque m�s relevantes, ya que no todos aplican en todos los casos. "
        prompt = prompt & "2. Redacta una descripci�n de la amenaza que incluya posibles escenarios de ataque, comenzando con la frase: -Un atacante podr�a...-. "
        prompt = prompt & "3. No es necesario proporcionar un ejemplo para cada vector de ataque. Selecciona el m�s realista o probable. "
        prompt = prompt & "4. En las vi�etas de vectores de ataque, menciona la probabilidad de ocurrencia, incluso si es baja. "
        
        prompt = prompt & "5. El contexto es un an�lisis de vulnerabilidades de infraestructura. Vulnerabilidad detectada mediante escaneo e interacciones. Formato de respuesta: � Responde en una tabla de dos columnas. � Para cada vulnerabilidad, redacta un p�rrafo descriptivo en la primera columna (m�nimo 75 palabras). � En la segunda columna, lista los vectores de ataque con vi�etas (usando guiones - ). � No uses HTML, solo texto plano. Ejemplo de estructura: Descripci�n de la amenaza   Vectores de ataque Un atacante podr�a explotar esta vulnerabilidad para acceder informaci�n"
        prompt = prompt & " confidencial. Esta amenaza es particularmente cr�tica en sistemas donde los controles de seguridad son menos estrictos."
        prompt = prompt & " Un escenario probable incluye...   - Malware: Un malware podr�a ser utilizado para... (probabilidad media). - Usuario malintencionado: Un empleado con acceso podr�a... (probabilidad baja). - Delincuente cibern�tico: Un atacante externo podr�a... (probabilidad alta). ES MUY IMPORTANTE QUE PARA LOS VECTORES DE ATAQUE DE LA AMENAZA USES GUIONES MEDIOS COMO VI�ETAS DENTRO DE LAS CELDAS. "
        
        prompt = prompt & "Explica los escenarios que consideres necesarios seg�n la naturaleza de la vulnerabilidad." & vbCrLf & Chr(10)
        prompt = prompt & "SOLO DOS columnas: Nombre (o descripci�n breve de la vulnerabilidad) y Amenaza/Vectores."
        prompt = prompt & " " & vbCrLf & Chr(10) & listaVulnerabilidades & vbCrLf & Chr(10)

        
        CopiarAlPortapapeles prompt

        
        MsgBox "El prompt para amenaza de vulnerabilidad ha sido copiado al portapapeles.", vbInformation, "Prompt Generado"
    Else
        MsgBox "No se encontraron valores en la selecci�n.", vbExclamation, "Sin Datos"
    End If
End Sub


Sub CYB070_PrepararPromptDesdeSeleccion_PropuestaRemediacionVuln_General()
    Dim celda As Range
    Dim listaVulnerabilidades As String
    Dim prompt As String

    
    listaVulnerabilidades = ""

    
    If TypeName(Selection) <> "Range" Then
        MsgBox "Por favor, selecciona un rango de celdas.", vbExclamation, "Selecci�n Inv�lida"
        Exit Sub
    End If

    
    For Each celda In Selection
        If Not IsEmpty(celda.value) Then
            If listaVulnerabilidades <> "" Then
                listaVulnerabilidades = listaVulnerabilidades & "," & vbCrLf
            End If
            listaVulnerabilidades = listaVulnerabilidades & Trim(CStr(celda.value))
        End If
    Next celda

    
    If listaVulnerabilidades <> "" Then
        prompt = "Hola, por favor Redacta como un pentester un p�rrafo t�cnico de propuesta de remediaci�n que comience con la frase: -Se recomienda...-. Incluye tantos detalles puntuales para la remediaci�n como sea posible, por ejemplo: nombres de soluciones que funcionan como protecci�n, controles de seguridad espec�ficos, dispositivos, configuraciones, buenas pr�cticas. Menciona de manera puntual qu� se podr�a hacer para que el encargado del sistema o activo pueda saber c�mo remediar. La respuesta debe tener SOLO un p�rrafo breve de introducci�n por vulnerabilidad y luego vi�etas (usando guiones -) para los puntos de la propuesta de remediaci�n. Responde para la siguiente lista de vulnerabilidades en FORMATO TABLA DE DOS COLUMNAS. SIEMPRE COMIENZA CON -Se recomienda...- TEXTO AMPLIO (m�s de 80 palabras por remediaci�n), aplicable a diversos casos, mencionando tecnolog�as, lenguajes o frameworks si aplica."
        
        prompt = prompt & " Menciona solo soluciones y pr�cticas corporativas/profesionales."
        prompt = prompt & " Solo dos columnas: Nombre (o descripci�n breve de la vulnerabilidad) y Propuesta de Remediaci�n."
        prompt = prompt & " " & vbCrLf & Chr(10) & listaVulnerabilidades & vbCrLf & Chr(10)

        
        CopiarAlPortapapeles prompt

        
        MsgBox "El prompt para propuesta de remediaci�n ha sido copiado al portapapeles.", vbInformation, "Prompt Generado"
    Else
        MsgBox "No se encontraron valores en la selecci�n.", vbExclamation, "Sin Datos"
    End If
End Sub

Sub CYB071_PrepararPromptDesdeSeleccion_DescripcionVuln_EnVPN()
    Dim celda As Range
    Dim listaVulnerabilidades As String
    Dim prompt As String
    
    
    listaVulnerabilidades = ""
    
    
    For Each celda In Selection
        If Not IsEmpty(celda.value) Then
            If listaVulnerabilidades <> "" Then
                listaVulnerabilidades = listaVulnerabilidades & "," & vbCrLf
            End If
            listaVulnerabilidades = listaVulnerabilidades & celda.value
        End If
    Next celda
    
    
    If listaVulnerabilidades <> "" Then
        prompt = "Redacta un p?rrafo t?cnico breve y conciso que describa la vulnerabilidad detectada, comenzando con la frase: -El sistema...-. Explica en qu? consiste la debilidad de seguridad de manera t?cnica. No incluyas escenarios de explotaci?n, ya que eso corresponde a otro campo. No describas c?mo se explota, solo en qu? consiste el problema. No menciones el nombre exacto de la vulnerabilidad; utiliza expresiones similares. Vulnerabilidades a describi en formato tabla: "
        prompt = prompt & " " & vbCrLf & Chr(10) & listaVulnerabilidades & vbCrLf & Chr(10)
        prompt = prompt & "Contexto de an?lisis: An?lisis de vulnerabilidades de infraestructura a partir conexion a red VPN. Vulnerabilidad detectada mediante escaneo e interacciones desde red privada VPN."

        
        
        CopiarAlPortapapeles prompt
        
        
        MsgBox "El prompt ha sido copiado al portapapeles.", vbInformation, "Prompt Generado"
    Else
        MsgBox "No se encontraron valores en la selecci?n.", vbExclamation, "Sin Datos"
    End If
End Sub



Sub CYB072_PrepararPromptDesdeSeleccion_AmenazaVuln_EnVPN()
    Dim celda As Range
    Dim listaVulnerabilidades As String
    Dim prompt As String
    
    
    listaVulnerabilidades = ""
    
    
    For Each celda In Selection
        If Not IsEmpty(celda.value) Then
            If listaVulnerabilidades <> "" Then
                listaVulnerabilidades = listaVulnerabilidades & "," & vbCrLf
            End If
            listaVulnerabilidades = listaVulnerabilidades & celda.value
        End If
    Next celda
    
    
    If listaVulnerabilidades <> "" Then
        
        prompt = "Hola, por favor considera el siguiente ejemplo: Un atacante podr�a (obtener, realizar, ejecutar, visualizar, identificar, listar)... Algunos posibles vectores de ataque adicionales asociados con esta amenaza incluyen (pueden ser algunos de los siguientes, pero no necesariamente todos): �  Malware: Un malware dise�ado para automatizar intentos de fuerza bruta podr�a explotar la vulnerabilidad para... �    Usuario malintencionado: Un usuario dentro de la red local con conocimiento de la vulnerabilidad podr�a aprovecharla para... �    Personal interno: Un empleado con acceso y conocimientos t?cnicos podr�a, intencionalmente o por error,... �  Delincuente cibern?tico: Un atacante externo en busca de vulnerabilidades podr�a intentar explotar esta debilidad para... Instrucciones adicionales: 1.   Pregunta si el sistema es interno o externo para determinar los vectores de ataque m?s relevantes, ya que no todos aplican en todos los casos. "
        prompt = prompt & "2.   Redacta una descripci?n de la amenaza que incluya posibles escenarios de ataque, comenzando con la frase: -Un atacante podr�a...-. 3.    No es necesario proporcionar un ejemplo para cada vector de ataque. Selecciona el m?s realista o probable. 4.   En las vi�etas de vectores de ataque, menciona la probabilidad de ocurrencia, incluso si es baja. 5.    El contexto es un an?lisis de vulnerabilidades de infraestructura desde una red privada, espec�ficamente: -Vulnerabilidad detectada mediante escaneo e interacciones desde red privada-. Formato de respuesta: �    Responde en una tabla de dos columnas. �    Para cada vulnerabilidad, redacta un p?rrafo descriptivo en la primera columna (m�nimo 75 palabras). �  En la segunda columna, lista los vectores de ataque con vi�etas (usando guiones  - ). � No uses HTML, solo texto plano. Ejemplo de estructura: Descripci?n de la amenaza    Vectores de ataque Un atacante podr�a explotar esta vulnerabilidad para acceder a inf _
maci?n"
        prompt = prompt & "confidencial Esta amenaza es particularmente cr�tica en sistemas internos donde los controles de seguridad son menos estrictos."
        prompt = prompt & "Un escenario probable incluye...    - Malware: Un malware podr�a ser utilizado para... (probabilidad media). - Usuario malintencionado: Un empleado con acceso podr�a... (probabilidad baja). - Delincuente cibern?tico: Un atacante externo podr�a... (probabilidad alta). ES MUY IMPORANTE QU EPAR ALOS VECTORI DE ATQUE DE LA MENAZA USSES GUIONE S MEDIOS COMO VI�ETAS DENTRO DE ALS CELDAS"
        prompt = prompt & "Se detecto mediante VPN red privada, pero explica los escenarios que consideres necesarios" & vbCrLf & Chr(10)
        prompt = prompt & "SOLO DOS columnas, nombre y amenaza"
        prompt = prompt & " " & vbCrLf & Chr(10) & listaVulnerabilidades & vbCrLf & Chr(10)
          
        
        CopiarAlPortapapeles prompt
        
        
        MsgBox "El prompt ha sido copiado al portapapeles.", vbInformation, "Prompt Generado"
    Else
        MsgBox "No se encontraron valores en la selecci?n.", vbExclamation, "Sin Datos"
    End If
End Sub


Sub CYB073_PrepararPromptDesdeSeleccion_PropuestaRemediacionVuln_EnVPN()
    Dim celda As Range
    Dim listaVulnerabilidades As String
    Dim prompt As String
    
    
    listaVulnerabilidades = ""
    
    
    For Each celda In Selection
        If Not IsEmpty(celda.value) Then
            If listaVulnerabilidades <> "" Then
                listaVulnerabilidades = listaVulnerabilidades & "," & vbCrLf
            End If
            listaVulnerabilidades = listaVulnerabilidades & celda.value
        End If
    Next celda
    
    
    If listaVulnerabilidades <> "" Then
        
        prompt = "Hola, por favor Redacta como un pentester un p?rrafo t?cnico de propuesta de remediaci?n que comience con la frase: -Se recomienda...-. Incluye tantos detalles puntuales par remaicion por jemplo nombre de soluciones que funeriona como poroteccon o controles de sguridad, disoitivos, practicas, menciona de manera m?s puntual que se pod�ria hacer para que le encargado del sistema, activo, pueda saber como remediari La respueste debe tener SOLO parrafo breve introducci?n y luego vi�etas para lso putnos de propueta de remedaicion Responde para la siguientel lista de vulnerabilidades en FORMATO TABLA DOS COLUMNAS, , SIEMPRE COMIENZAX CON -Se recomienda...- TEXTO AMPLIDO MAS DE 80 palabras, atalmente aplicable a muchos casos, lengauje so framworks"
           prompt = prompt & "Se detecto mediante VPN red privada, pero explica los escenarios que consideres necesarios" & vbCrLf & Chr(10)
          prompt = prompt & "Menciona solo soluciones corporativas"
          prompt = prompt & "Solo dos columnas, nombre y propuesat remediacion"
        prompt = prompt & " " & vbCrLf & Chr(10) & listaVulnerabilidades & vbCrLf & Chr(10)
          
        
        CopiarAlPortapapeles prompt
        
        
        MsgBox "El prompt ha sido copiado al portapapeles.", vbInformation, "Prompt Generado"
    Else
        MsgBox "No se encontraron valores en la selecci?n.", vbExclamation, "Sin Datos"
    End If
End Sub


Sub CYB074_PrepararPromptDesdeSeleccion_DescripcionVuln_EnRedPrivada()
    Dim celda As Range
    Dim listaVulnerabilidades As String
    Dim prompt As String
    
    
    listaVulnerabilidades = ""
    
    
    For Each celda In Selection
        If Not IsEmpty(celda.value) Then
            If listaVulnerabilidades <> "" Then
                listaVulnerabilidades = listaVulnerabilidades & "," & vbCrLf
            End If
            listaVulnerabilidades = listaVulnerabilidades & celda.value
        End If
    Next celda
    
    
    If listaVulnerabilidades <> "" Then
        prompt = "Redacta un p?rrafo t?cnico breve y conciso que describa la vulnerabilidad detectada, comenzando con la frase: -El sistema...-. Explica en qu? consiste la debilidad de seguridad de manera t?cnica. No incluyas escenarios de explotaci?n, ya que eso corresponde a otro campo. No describas c?mo se explota, solo en qu? consiste el problema. No menciones el nombre exacto de la vulnerabilidad; utiliza expresiones similares. Vulnerabilidades a describi en formato tabla: "
        prompt = prompt & " " & vbCrLf & Chr(10) & listaVulnerabilidades & vbCrLf & Chr(10)
        prompt = prompt & "Contexto de an?lisis: An?lisis de vulnerabilidades de infraestructura a partir conexion a red en red privada en sitio. Vulnerabilidad detectada mediante escaneo e interacciones desde red privada en red privada en sitio."

        
        
        CopiarAlPortapapeles prompt
        
        
        MsgBox "El prompt ha sido copiado al portapapeles.", vbInformation, "Prompt Generado"
    Else
        MsgBox "No se encontraron valores en la selecci?n.", vbExclamation, "Sin Datos"
    End If
End Sub



Sub CYB075_PrepararPromptDesdeSeleccion_AmenazaVuln_EnRedPrivada()
    Dim celda As Range
    Dim listaVulnerabilidades As String
    Dim prompt As String
    
    
    listaVulnerabilidades = ""
    
    
    For Each celda In Selection
        If Not IsEmpty(celda.value) Then
            If listaVulnerabilidades <> "" Then
                listaVulnerabilidades = listaVulnerabilidades & "," & vbCrLf
            End If
            listaVulnerabilidades = listaVulnerabilidades & celda.value
        End If
    Next celda
    
    
    If listaVulnerabilidades <> "" Then
        
        prompt = "Hola, por favor considera el siguiente ejemplo: Un atacante podr�a (obtener, realizar, ejecutar, visualizar, identificar, listar)... Algunos posibles vectores de ataque adicionales asociados con esta amenaza incluyen (pueden ser algunos de los siguientes, pero no necesariamente todos): �  Malware: Un malware dise�ado para automatizar intentos de fuerza bruta podr�a explotar la vulnerabilidad para... �    Usuario malintencionado: Un usuario dentro de la red local con conocimiento de la vulnerabilidad podr�a aprovecharla para... �    Personal interno: Un empleado con acceso y conocimientos t?cnicos podr�a, intencionalmente o por error,... �  Delincuente cibern?tico: Un atacante externo en busca de vulnerabilidades podr�a intentar explotar esta debilidad para... Instrucciones adicionales: 1.   Pregunta si el sistema es interno o externo para determinar los vectores de ataque m?s relevantes, ya que no todos aplican en todos los casos. "
        prompt = prompt & "2.   Redacta una descripci?n de la amenaza que incluya posibles escenarios de ataque, comenzando con la frase: -Un atacante podr�a...-. 3.    No es necesario proporcionar un ejemplo para cada vector de ataque. Selecciona el m?s realista o probable. 4.   En las vi�etas de vectores de ataque, menciona la probabilidad de ocurrencia, incluso si es baja. 5.    El contexto es un an?lisis de vulnerabilidades de infraestructura desde una red privada, espec�ficamente: -Vulnerabilidad detectada mediante escaneo e interacciones desde red privada-. Formato de respuesta: �    Responde en una tabla de dos columnas. �    Para cada vulnerabilidad, redacta un p?rrafo descriptivo en la primera columna (m�nimo 75 palabras). �  En la segunda columna, lista los vectores de ataque con vi�etas (usando guiones  - ). � No uses HTML, solo texto plano. Ejemplo de estructura: Descripci?n de la amenaza    Vectores de ataque Un atacante podr�a explotar esta vulnerabilidad para acceder a inf _
maci?n"
        prompt = prompt & "confidencial Esta amenaza es particularmente cr�tica en sistemas internos donde los controles de seguridad son menos estrictos."
        prompt = prompt & "Un escenario probable incluye...    - Malware: Un malware podr�a ser utilizado para... (probabilidad media). - Usuario malintencionado: Un empleado con acceso podr�a... (probabilidad baja). - Delincuente cibern?tico: Un atacante externo podr�a... (probabilidad alta). ES MUY IMPORANTE QU EPAR ALOS VECTORI DE ATQUE DE LA MENAZA USSES GUIONE S MEDIOS COMO VI�ETAS DENTRO DE ALS CELDAS"
        prompt = prompt & "Se detecto mediante en red privada en sitio red privada, pero explica los escenarios que consideres necesarios" & vbCrLf & Chr(10)
        prompt = prompt & "SOLO DOS columnas, nombre y amenaza"
        prompt = prompt & " " & vbCrLf & Chr(10) & listaVulnerabilidades & vbCrLf & Chr(10)
          
        
        CopiarAlPortapapeles prompt
        
        
        MsgBox "El prompt ha sido copiado al portapapeles.", vbInformation, "Prompt Generado"
    Else
        MsgBox "No se encontraron valores en la selecci?n.", vbExclamation, "Sin Datos"
    End If
End Sub


Sub CYB076_PrepararPromptDesdeSeleccion_PropuestaRemediacionVuln_EnRedPrivada()
    Dim celda As Range
    Dim listaVulnerabilidades As String
    Dim prompt As String
    
    
    listaVulnerabilidades = ""
    
    
    For Each celda In Selection
        If Not IsEmpty(celda.value) Then
            If listaVulnerabilidades <> "" Then
                listaVulnerabilidades = listaVulnerabilidades & "," & vbCrLf
            End If
            listaVulnerabilidades = listaVulnerabilidades & celda.value
        End If
    Next celda
    
    
    If listaVulnerabilidades <> "" Then
        
        prompt = "Hola, por favor Redacta como un pentester un p?rrafo t?cnico de propuesta de remediaci?n que comience con la frase: -Se recomienda...-. Incluye tantos detalles puntuales par remaicion por jemplo nombre de soluciones que funeriona como poroteccon o controles de sguridad, disoitivos, practicas, menciona de manera m?s puntual que se pod�ria hacer para que le encargado del sistema, activo, pueda saber como remediari La respueste debe tener SOLO parrafo breve introducci?n y luego vi�etas para lso putnos de propueta de remedaicion Responde para la siguientel lista de vulnerabilidades en FORMATO TABLA DOS COLUMNAS, , SIEMPRE COMIENZAX CON -Se recomienda...- TEXTO AMPLIDO MAS DE 80 palabras, atalmente aplicable a muchos casos, lengauje so framworks"
           prompt = prompt & "Se detecto mediante en red privada en sitio red privada, pero explica los escenarios que consideres necesarios" & vbCrLf & Chr(10)
          prompt = prompt & "Menciona solo soluciones corporativas"
          prompt = prompt & "Solo dos columnas, nombre y propuesat remediacion"
        prompt = prompt & " " & vbCrLf & Chr(10) & listaVulnerabilidades & vbCrLf & Chr(10)
          
        
        CopiarAlPortapapeles prompt
        
        
        MsgBox "El prompt ha sido copiado al portapapeles.", vbInformation, "Prompt Generado"
    Else
        MsgBox "No se encontraron valores en la selecci?n.", vbExclamation, "Sin Datos"
    End If
End Sub

Sub CYB0761_PrepararPromptDesdeSeleccion_ExplicacionTecnicaVuln_EnRedPrivada()
    Dim celda As Range
    Dim listaVulnerabilidades As String
    Dim prompt As String
    
    
    listaVulnerabilidades = ""
    
    
    For Each celda In Selection
        If Not IsEmpty(celda.value) Then
            If listaVulnerabilidades <> "" Then
                listaVulnerabilidades = listaVulnerabilidades & "," & vbCrLf
            End If
            listaVulnerabilidades = listaVulnerabilidades & celda.value
        End If
    Next celda
    
    
    If listaVulnerabilidades <> "" Then
       
prompt = "Hola, por favor en una tabla, solo dos columnas: vulnerabilidad y explicaci?n t?cnica. "
prompt = prompt & "Para cada una de estas vulnerabilidades redacta un p?rrafo de explicaci?n t?cnica que contenga un ejemplo y "
prompt = prompt & "una conclusi?n breve, concisa y convincente desde la perspectiva de pentesting. "
prompt = prompt & "Inicia la explicaci?n con el texto -En un escenario...- [t�pico / com�n / poco probable]. "
prompt = prompt & "De ser posible, agrega c?digo de ejemplo para comprender este tipo de vulnerabilidad. "
prompt = prompt & "El c?digo debe ser �til, no seas escaso en detalles. "
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
prompt = prompt & "QUIERO UNA TABLA CON BUEN FORMATO EN LAS CELDAS, SALTOS DE L�NEA APROPIADOS. "
prompt = prompt & "M�S DE 125 CARACTERES. "
prompt = prompt & "NO PONGAS TODO EN UN SOLO P�RRAFO, USA SALTOS DE L�NEA DENTRO DE LAS CELDAS DE EXPLICACI�N PARA QUE SEA LEGIBLE. "
        prompt = prompt & " " & vbCrLf & Chr(10) & listaVulnerabilidades & vbCrLf & Chr(10)
        
        
        CopiarAlPortapapeles prompt
        
        
        MsgBox "El prompt ha sido copiado al portapapeles.", vbInformation, "Prompt Generado"
    Else
        MsgBox "No se encontraron valores en la selecci?n.", vbExclamation, "Sin Datos"
    End If
End Sub

Sub CYB0762_PrepararPromptDesdeSeleccion_VectorCVSSVuln_EnRedPrivada()
    Dim celda As Range
    Dim listaVulnerabilidades As String
    Dim prompt As String
    
    
    listaVulnerabilidades = ""
    
    
    For Each celda In Selection
        If Not IsEmpty(celda.value) Then
            If listaVulnerabilidades <> "" Then
                listaVulnerabilidades = listaVulnerabilidades & "," & vbCrLf
            End If
            listaVulnerabilidades = listaVulnerabilidades & celda.value
        End If
    Next celda
    
    
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
        prompt = prompt & "Por favor eval�e la severidad con base en los criterios anteriores, siendo exigente y estricto al asignar la severidad del CVSS. No asigne impactos altos a menos que haya evidencia de que se pueda explotar directamente y afecte de manera significativa." & vbCrLf
        prompt = prompt & "SOLO RESPONDE CADENAS VECTOR COMPLETAS EJEMPLO CVSS:4.0/AV:A/AC:L/AT:N/PR:N/UI:N/VC:N/VI:N/VA:N/SC:N/SI:N/SA:N VULNEBILIDAD, CVSS" & vbCrLf
        prompt = prompt & "VULNEBILIDAD, CVSS" & vbCrLf
        prompt = prompt & " " & vbCrLf & Chr(10) & listaVulnerabilidades & vbCrLf & Chr(10)
        
        
        CopiarAlPortapapeles prompt
        
        
        MsgBox "El prompt ha sido copiado al portapapeles.", vbInformation, "Prompt Generado"
    Else
        MsgBox "No se encontraron valores en la selecci?n.", vbExclamation, "Sin Datos"
    End If
End Sub

Sub CYB077_PrepararPromptDesdeSeleccion_DescripcionVuln_DesdeInternet()
    Dim celda As Range
    Dim listaVulnerabilidades As String
    Dim prompt As String
    
    
    listaVulnerabilidades = ""
    
    
    For Each celda In Selection
        If Not IsEmpty(celda.value) Then
            If listaVulnerabilidades <> "" Then
                listaVulnerabilidades = listaVulnerabilidades & "," & vbCrLf
            End If
            listaVulnerabilidades = listaVulnerabilidades & celda.value
        End If
    Next celda
    
    
    If listaVulnerabilidades <> "" Then
        prompt = "Redacta un p?rrafo t?cnico breve y conciso que describa la vulnerabilidad detectada, comenzando con la frase: -El sistema...-. Explica en qu? consiste la debilidad de seguridad de manera t?cnica. No incluyas escenarios de explotaci?n, ya que eso corresponde a otro campo. No describas c?mo se explota, solo en qu? consiste el problema. No menciones el nombre exacto de la vulnerabilidad; utiliza expresiones similares. Vulnerabilidades a describi en formato tabla: "
        prompt = prompt & " " & vbCrLf & Chr(10) & listaVulnerabilidades & vbCrLf & Chr(10)
        prompt = prompt & "Contexto de an?lisis: An?lisis de vulnerabilidades de infraestructura a partir conexion a red en desde internet en sitio. Vulnerabilidad detectada mediante escaneo e interacciones desde desde internet en desde internet en sitio."

        
        
        CopiarAlPortapapeles prompt
        
        
        MsgBox "El prompt ha sido copiado al portapapeles.", vbInformation, "Prompt Generado"
    Else
        MsgBox "No se encontraron valores en la selecci?n.", vbExclamation, "Sin Datos"
    End If
End Sub



Sub CYB078_PrepararPromptDesdeSeleccion_AmenazaVuln_DesdeInternet()
    Dim celda As Range
    Dim listaVulnerabilidades As String
    Dim prompt As String
    
    
    listaVulnerabilidades = ""
    
    
    For Each celda In Selection
        If Not IsEmpty(celda.value) Then
            If listaVulnerabilidades <> "" Then
                listaVulnerabilidades = listaVulnerabilidades & "," & vbCrLf
            End If
            listaVulnerabilidades = listaVulnerabilidades & celda.value
        End If
    Next celda
    
    
    If listaVulnerabilidades <> "" Then
        
        prompt = "Hola, por favor considera el siguiente ejemplo: Un atacante podr�a (obtener, realizar, ejecutar, visualizar, identificar, listar)... Algunos posibles vectores de ataque adicionales asociados con esta amenaza incluyen (pueden ser algunos de los siguientes, pero no necesariamente todos): �  Malware: Un malware dise�ado para automatizar intentos de fuerza bruta podr�a explotar la vulnerabilidad para... �    Usuario malintencionado: Un usuario dentro de la red local con conocimiento de la vulnerabilidad podr�a aprovecharla para... �    Personal interno: Un empleado con acceso y conocimientos t?cnicos podr�a, intencionalmente o por error,... �  Delincuente cibern?tico: Un atacante externo en busca de vulnerabilidades podr�a intentar explotar esta debilidad para... Instrucciones adicionales: 1.   Pregunta si el sistema es interno o externo para determinar los vectores de ataque m?s relevantes, ya que no todos aplican en todos los casos. "
        prompt = prompt & "2.   Redacta una descripci?n de la amenaza que incluya posibles escenarios de ataque, comenzando con la frase: -Un atacante podr�a...-. 3.    No es necesario proporcionar un ejemplo para cada vector de ataque. Selecciona el m?s realista o probable. 4.   En las vi�etas de vectores de ataque, menciona la probabilidad de ocurrencia, incluso si es baja. 5.    El contexto es un an?lisis de vulnerabilidades de infraestructura desde una desde internet, espec�ficamente: -Vulnerabilidad detectada mediante escaneo e interacciones desde desde internet-. Formato de respuesta: �    Responde en una tabla de dos columnas. �    Para cada vulnerabilidad, redacta un p?rrafo descriptivo en la primera columna (m�nimo 75 palabras). �  En la segunda columna, lista los vectores de ataque con vi�etas (usando guiones  - ). � No uses HTML, solo texto plano. Ejemplo de estructura: Descripci?n de la amenaza    Vectores de ataque Un atacante podr�a explotar esta vulnerabilidad para acceder informacion """
        prompt = prompt & "confidencial Esta amenaza es particularmente cr�tica en sistemas internos donde los controles de seguridad son menos estrictos."
        prompt = prompt & "Un escenario probable incluye...    - Malware: Un malware podr�a ser utilizado para... (probabilidad media). - Usuario malintencionado: Un empleado con acceso podr�a... (probabilidad baja). - Delincuente cibern?tico: Un atacante externo podr�a... (probabilidad alta). ES MUY IMPORANTE QU EPAR ALOS VECTORI DE ATQUE DE LA MENAZA USSES GUIONE S MEDIOS COMO VI�ETAS DENTRO DE ALS CELDAS"
        prompt = prompt & "Se detecto mediante en desde internet en sitio desde internet, pero explica los escenarios que consideres necesarios" & vbCrLf & Chr(10)
        prompt = prompt & "SOLO DOS columnas, nombre y amenaza"
        prompt = prompt & " " & vbCrLf & Chr(10) & listaVulnerabilidades & vbCrLf & Chr(10)
          
        
        CopiarAlPortapapeles prompt
        
        
        MsgBox "El prompt ha sido copiado al portapapeles.", vbInformation, "Prompt Generado"
    Else
        MsgBox "No se encontraron valores en la selecci?n.", vbExclamation, "Sin Datos"
    End If
End Sub


Sub CYB079_PrepararPromptDesdeSeleccion_PropuestaRemediacionVuln_DesdeInternet()
    Dim celda As Range
    Dim listaVulnerabilidades As String
    Dim prompt As String
    
    
    listaVulnerabilidades = ""
    
    
    For Each celda In Selection
        If Not IsEmpty(celda.value) Then
            If listaVulnerabilidades <> "" Then
                listaVulnerabilidades = listaVulnerabilidades & "," & vbCrLf
            End If
            listaVulnerabilidades = listaVulnerabilidades & celda.value
        End If
    Next celda
    
    
    If listaVulnerabilidades <> "" Then
        
        prompt = "Hola, por favor Redacta como un pentester un p?rrafo t?cnico de propuesta de remediaci?n que comience con la frase: -Se recomienda...-. Incluye tantos detalles puntuales par remaicion por jemplo nombre de soluciones que funeriona como poroteccon o controles de sguridad, disoitivos, practicas, menciona de manera m?s puntual que se pod�ria hacer para que le encargado del sistema, activo, pueda saber como remediari La respueste debe tener SOLO parrafo breve introducci?n y luego vi�etas para lso putnos de propueta de remedaicion Responde para la siguientel lista de vulnerabilidades en FORMATO TABLA DOS COLUMNAS, , SIEMPRE COMIENZAX CON -Se recomienda...- TEXTO AMPLIDO MAS DE 80 palabras, atalmente aplicable a muchos casos, lengauje so framworks"
           prompt = prompt & "Se detecto mediante en desde internet en sitio desde internet, pero explica los escenarios que consideres necesarios" & vbCrLf & Chr(10)
          prompt = prompt & "Menciona solo soluciones corporativas"
          prompt = prompt & "Solo dos columnas, nombre y propuesat remediacion"
        prompt = prompt & " " & vbCrLf & Chr(10) & listaVulnerabilidades & vbCrLf & Chr(10)
          
        
        CopiarAlPortapapeles prompt
        
        
        MsgBox "El prompt ha sido copiado al portapapeles.", vbInformation, "Prompt Generado"
    Else
        MsgBox "No se encontraron valores en la selecci?n.", vbExclamation, "Sin Datos"
    End If
End Sub

Sub CYB080_PreparePromptFromSelection_DescripcionVuln_FromCode()
    Dim celda As Range
    Dim listaVulnerabilidades As String
    Dim prompt As String
    
    
    listaVulnerabilidades = ""
    
    
    For Each celda In Selection
        If Not IsEmpty(celda.value) Then
            If listaVulnerabilidades <> "" Then
                listaVulnerabilidades = listaVulnerabilidades & "," & vbCrLf
            End If
            listaVulnerabilidades = listaVulnerabilidades & celda.value
        End If
    Next celda
    
    
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

        
        CopiarAlPortapapeles prompt
        
        
        MsgBox "El prompt ha sido copiado al portapapeles.", vbInformation, "Prompt Generado"
    Else
        MsgBox "No se encontraron valores en la selecci?n.", vbExclamation, "Sin Datos"
    End If
End Sub

Sub CYB081_PreparePromptFromSelection_AmenazaVuln_FromCode()
    Dim celda As Range
    Dim listaVulnerabilidades As String
    Dim prompt As String
    
    
    listaVulnerabilidades = ""
    
    
    For Each celda In Selection
        If Not IsEmpty(celda.value) Then
            If listaVulnerabilidades <> "" Then
                listaVulnerabilidades = listaVulnerabilidades & "," & vbCrLf
            End If
            listaVulnerabilidades = listaVulnerabilidades & celda.value
        End If
    Next celda
    
    
    If listaVulnerabilidades <> "" Then
        prompt = ""
        prompt = prompt & "Hola, por favor considera el siguiente ejemplo: " & vbCrLf
        prompt = prompt & "Un atacante podr�a (inyectar, manipular, filtrar, exponer, escalar privilegios)..." & vbCrLf
        prompt = prompt & "Algunos posibles vectores de ataque adicionales asociados con esta amenaza incluyen (pueden ser algunos de los siguientes, pero no necesariamente todos):" & vbCrLf
        prompt = prompt & "�  Inyecci?n de c?digo: Una entrada no validada en el c?digo fuente podr�a permitir inyecci?n de comandos..." & vbCrLf
        prompt = prompt & "�  Exposici?n de informaci?n sensible: Una mala gesti?n de credenciales en el c?digo podr�a revelar secretos..." & vbCrLf
        prompt = prompt & "�  Elevaci?n de privilegios: Una funci?n mal dise�ada podr�a permitir a un usuario ejecutar acciones con m?s permisos de los necesarios..." & vbCrLf
        prompt = prompt & "�  Manipulaci?n de datos: Un atacante podr�a modificar par?metros dentro del c?digo para alterar la l?gica de la aplicaci?n..." & vbCrLf
        prompt = prompt & vbCrLf
        prompt = prompt & "Instrucciones adicionales:" & vbCrLf
        prompt = prompt & "1. Pregunta si el c?digo pertenece a una aplicaci?n interna o externa para determinar los vectores de ataque m?s relevantes." & vbCrLf
        prompt = prompt & "2. Redacta una descripci?n de la amenaza que incluya posibles escenarios de ataque, comenzando con la frase: -Un atacante podr�a...-." & vbCrLf
        prompt = prompt & "3. No es necesario proporcionar un ejemplo para cada vector de ataque. Selecciona el m?s realista o probable." & vbCrLf
        prompt = prompt & "4. En las vi�etas de vectores de ataque, menciona la probabilidad de ocurrencia, incluso si es baja." & vbCrLf
        prompt = prompt & "5. El contexto es un an?lisis de c?digo fuente mediante herramientas de an?lisis est?tico (SAST), sin ejecutar la aplicaci?n." & vbCrLf
        prompt = prompt & vbCrLf
        prompt = prompt & "Formato de respuesta:" & vbCrLf
        prompt = prompt & "� Responde en una tabla de dos columnas." & vbCrLf
        prompt = prompt & "� Para cada vulnerabilidad, redacta un p?rrafo descriptivo en la primera columna (m�nimo 75 palabras)." & vbCrLf
        prompt = prompt & "� En la segunda columna, lista los vectores de ataque con vi�etas (usando guiones - )." & vbCrLf
        prompt = prompt & "� No uses HTML, solo texto plano." & vbCrLf
        prompt = prompt & vbCrLf
        prompt = prompt & "Ejemplo de estructura:" & vbCrLf
        prompt = prompt & "Descripci?n de la amenaza    Vectores de ataque" & vbCrLf
        prompt = prompt & "Un atacante podr�a explotar esta vulnerabilidad en el c?digo para ejecutar comandos arbitrarios..." & vbCrLf
        prompt = prompt & "    - Inyecci?n de c?digo: Un usuario malintencionado podr�a insertar c?digo malicioso... (probabilidad media)." & vbCrLf
        prompt = prompt & "    - Exposici?n de informaci?n: Credenciales en el c?digo podr�an filtrarse... (probabilidad alta)." & vbCrLf
        prompt = prompt & vbCrLf
        prompt = prompt & "ES MUY IMPORTANTE QUE PARA LOS VECTORES DE ATAQUE DE LA AMENAZA USES GUIONES MEDIOS COMO VI�ETAS DENTRO DE LAS CELDAS." & vbCrLf
        prompt = prompt & "An?lisis realizado mediante herramientas SAST en c?digo fuente est?tico sin ejecuci?n." & vbCrLf
        prompt = prompt & "SOLO DOS COLUMNAS: NOMBRE Y AMENAZA." & vbCrLf
        prompt = prompt & vbCrLf & listaVulnerabilidades & vbCrLf & vbCrLf
          
        
        CopiarAlPortapapeles prompt
        
        
        MsgBox "El prompt ha sido copiado al portapapeles.", vbInformation, "Prompt Generado"
    Else
        MsgBox "No se encontraron valores en la selecci?n.", vbExclamation, "Sin Datos"
    End If
End Sub


Sub CYB082_PreparePromptFromSelection_PropuestaRemediacionVuln_FromCode()
    Dim celda As Range
    Dim listaVulnerabilidades As String
    Dim prompt As String
    
    
    listaVulnerabilidades = ""
    
    
    For Each celda In Selection
        If Not IsEmpty(celda.value) Then
            If listaVulnerabilidades <> "" Then
                listaVulnerabilidades = listaVulnerabilidades & "," & vbCrLf
            End If
            listaVulnerabilidades = listaVulnerabilidades & celda.value
        End If
    Next celda
    
    
    If listaVulnerabilidades <> "" Then
        prompt = "Hola, por favor redacta como un pentester un p?rrafo t?cnico de propuesta de remediaci?n que comience con la frase: -Se recomienda...-."
        prompt = prompt & " Incluye tantos detalles puntuales como sea posible, mencionando soluciones espec�ficas, controles de seguridad, dispositivos y pr?cticas recomendadas."
        prompt = prompt & " Proporciona informaci?n clara para que el encargado del sistema o activo sepa exactamente c?mo remediarlo."
        prompt = prompt & " La respuesta debe contener un p?rrafo breve de introducci?n seguido de vi�etas con los puntos de la propuesta de remediaci?n."
        prompt = prompt & " Formato de respuesta: una tabla de dos columnas."
        prompt = prompt & " Siempre comienza con -Se recomienda...-."
        prompt = prompt & " El texto debe ser amplio, con m?s de 80 palabras, aplicable a m�ltiples casos y en lenguaje t?cnico adecuado."
        prompt = prompt & " Se detect? mediante an?lisis desde internet en el sitio, pero explica los escenarios relevantes."
        prompt = prompt & " Menciona solo soluciones corporativas."
        prompt = prompt & " Solo dos columnas: nombre y propuesta de remediaci?n."
        prompt = prompt & " " & vbCrLf & Chr(10) & listaVulnerabilidades & vbCrLf & Chr(10)
        
        
        CopiarAlPortapapeles prompt
        
        
        MsgBox "El prompt ha sido copiado al portapapeles.", vbInformation, "Prompt Generado"
    Else
        MsgBox "No se encontraron valores en la selecci?n.", vbExclamation, "Sin Datos"
    End If
End Sub

Sub CYB082_PreparePromptFromSelection_MetodoDeteccion()
    Dim celda As Range
    Dim listaVulnerabilidades As String
    Dim prompt As String
    
    
    listaVulnerabilidades = ""
    
    
    For Each celda In Selection
        If Not IsEmpty(celda.value) Then
            If listaVulnerabilidades <> "" Then
                listaVulnerabilidades = listaVulnerabilidades & "," & vbCrLf
            End If
            listaVulnerabilidades = listaVulnerabilidades & celda.value
        End If
    Next celda
    
    
If listaVulnerabilidades <> "" Then
    prompt = "Redacta un p�rrafo t�cnico breve que describa c�mo se identific� una vulnerabilidad a partir de la salida de una herramienta de an�lisis de seguridad."
    prompt = prompt & " El p�rrafo debe seguir esta estructura:"
    prompt = prompt & " 1. Comienza indicando que se detect� la vulnerabilidad."
    prompt = prompt & " 2. Menciona la herramienta empleada y la t�cnica o enfoque utilizado (por ejemplo: escaneo de red, an�lisis de cabeceras, env�o de solicitudes HTTP, validaci�n de configuraciones, etc.)."
    prompt = prompt & " 3. Resume de manera concisa qu� analiza la herramienta o qu� tipo de fallo identifica, mencionando el servicio, tecnolog�a o componente afectado."
    prompt = prompt & " 4. Opcionalmente, incluye entre par�ntesis el nombre del script o m�dulo espec�fico usado."
    prompt = prompt & " Usa un lenguaje t�cnico adecuado y profesional."
    prompt = prompt & " Ejemplos:"
    prompt = prompt & " - Detectamos esta vulnerabilidad mediante una herramienta que env�a solicitudes HTTP del tipo GET (shcheck.py) y examina la presencia o ausencia de cabeceras que usa el navegador web para asegurar algunas interacciones con el usuario final."
    prompt = prompt & " - Usamos un esc�ner de red (Nessus) para detectar fallos en Apache que permiten la explotaci�n de vulnerabilidades conocidas que podr�an comprometer la disponibilidad y seguridad de los servicios web."
    prompt = prompt & " - La vulnerabilidad fue identificada utilizando una herramienta de escaneo automatizado (Acunetix), que eval�a configuraciones inseguras y comportamientos an�malos en aplicaciones web."
    prompt = prompt & " Herramientas t�picas a considerar: Acunetix, shcheck.py, sshcrik, Nessus, OpenVAS, Nexpose, sqlmap, entre otras."
    prompt = prompt & " Genera el p�rrafo bas�ndote en la siguiente salida de herramienta:"
    prompt = prompt & vbCrLf & Chr(10) & listaVulnerabilidades & vbCrLf & Chr(10)
       
        
        CopiarAlPortapapeles prompt
        
        
        MsgBox "El prompt ha sido copiado al portapapeles.", vbInformation, "Prompt Generado"
    Else
        MsgBox "No se encontraron valores en la selecci?n.", vbExclamation, "Sin Datos"
    End If
End Sub




Sub CYB083_Verificar_VectorCVSS4_0()
    Dim celda As Range
    Dim cvssString As String
    Dim url As String
    
    
    cvssString = ""
    
    
    For Each celda In Selection
        If Not IsEmpty(celda.value) Then
            If cvssString <> "" Then
                cvssString = cvssString & "/" & celda.value
            Else
                cvssString = celda.value
            End If
        End If
    Next celda
    
    
    If cvssString <> "" Then
        
        url = "https://www.first.org/cvss/calculator/4.0#" & cvssString
        
        
        CopiarAlPortapapeles url
        
        
        MsgBox "La URL ha sido copiada al portapapeles: " & vbCrLf & url, vbInformation, "URL Generada"
        
        
        ThisWorkbook.FollowHyperlink url
    Else
        MsgBox "No se encontraron valores en la selecci?n.", vbExclamation, "Sin Datos"
    End If
End Sub


Sub CopiarAlPortapapeles(text As String)
    
    Dim MSForms_DataObject As Object
    Set MSForms_DataObject = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    MSForms_DataObject.SetText text
    MSForms_DataObject.PutInClipboard
    Set MSForms_DataObject = Nothing
End Sub




Public Sub InsertarTextoMarkdownEnWordConFormato(WordApp As Object, WordDoc As Object, placeholder As String, markdownText As String, basePath As String, CustomStyle As Boolean, CustomStyleName As String)

    
    Const wdStory As Long = 6
    Const wdFindContinue As Long = 1
    Const wdReplaceOne As Long = 1
    Const wdReplaceAll As Long = 2
    Const wdCollapseEnd As Long = 0
    Const wdLineStyleSingle As Long = 1
    Const wdLineWidth050pt As Long = 4
    Const wdColorGray25 As Long = 14277081
    Const wdColorAutomatic As Long = -16777216
    Const wdAlignParagraphCenter As Long = 1
    Const wdAlignParagraphLeft As Long = 0
    Const wdListApplyToSelection As Long = 0
    Const wdListNumberStyleBullet As Long = 23
    Const wdContinueList As Long = 1
    Const wdRestartNumbering As Long = 0
    Const wdListNoNumbering As Long = 0
    Const wdFormatDocument As Long = 0
    Const wdNumberListNum As Long = 2
    Const wdBulletListNum As Long = 1
    
    Const wdBorderTop As Long = -1
    Const wdBorderLeft As Long = -2
    Const wdBorderBottom As Long = -3
    Const wdBorderRight As Long = -4
    Const wdBorderHorizontal As Long = -5
    Const wdBorderVertical As Long = -6

    
    Dim sel As Object
    Dim lines() As String
    Dim i As Long
    Dim lineText As String
    Dim trimmedLine As String
    Dim inCodeBlock As Boolean
    Dim fs As Object
    Dim fullImgPath As String
    Dim codeBlockStartRange As Object
    Dim currentListType As String
    Dim lastLineWasList As Boolean
    Dim hLevel As Integer
    Dim contentText As String
    Dim listMarkerPos As Integer
    Dim isPathAbsolute As Boolean
    Dim paraRange As Object

    
    On Error GoTo ErrorHandler

    
    If WordApp Is Nothing Then MsgBox "Error cr�tico: La variable "
    If WordDoc Is Nothing Then MsgBox "Error cr�tico: La variable "
    On Error Resume Next
    Dim testAppName As String: testAppName = WordApp.Name
    If Err.Number <> 0 Then MsgBox "Error cr�tico: La variable "
    Dim testDocName As String: testDocName = WordDoc.Name
    If Err.Number <> 0 Then MsgBox "Error cr�tico: La variable "
    On Error GoTo ErrorHandler

    
    On Error Resume Next
    Set fs = CreateObject("Scripting.FileSystemObject")
    If Err.Number <> 0 Or fs Is Nothing Then
        MsgBox "Error cr�tico: No se pudo crear el "
        Exit Sub
    End If
    On Error GoTo ErrorHandler

    
    On Error Resume Next
    WordDoc.Activate
    WordApp.Activate
    If Err.Number <> 0 Then
        Debug.Print "Advertencia: No se pudo activar la ventana de Word o el documento. Se continuar�, pero podr�a haber problemas si Word no est� visible/activo. Error: " & Err.Description
        Err.Clear
    End If
    On Error GoTo ErrorHandler

    Set sel = WordApp.Selection
    If sel Is Nothing Then MsgBox "Error cr�tico: No se pudo obtener el objeto "

    
    If Len(Trim(basePath)) > 0 Then
        If Right(basePath, 1) <> fs.GetStandardStream(1).Write(vbNullString) And Right(basePath, 1) <> "/" Then
            Dim pathSep As String
             On Error Resume Next
             pathSep = WordApp.PathSeparator
             If Err.Number <> 0 Then pathSep = "\"
             Err.Clear
             On Error GoTo ErrorHandler
            basePath = basePath & pathSep
             Debug.Print "BasePath con separador a�adido manualmente: "
         End If
         Debug.Print "BasePath final: "
    Else
        Debug.Print "No se proporcion� BasePath. Las rutas relativas de im�genes podr�an fallar si se procesan posteriormente."
    End If


    
    
    sel.HomeKey wdStory
    sel.Find.ClearFormatting
    sel.Find.Replacement.ClearFormatting
    With sel.Find
        .text = placeholder
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceOne
    End With


    If sel.Find.found Then
        Debug.Print "Placeholder "

        
        markdownText = Replace(markdownText, vbCrLf, vbLf)
        markdownText = Replace(markdownText, vbCr, vbLf)
        lines = Split(markdownText, vbLf)
        
        inCodeBlock = False
        currentListType = ""
        lastLineWasList = False
        Set codeBlockStartRange = Nothing

        
        On Error Resume Next
        If CustomStyle Then sel.Range.Style = WordDoc.Styles("Normal")
        If Err.Number <> 0 Then
            Debug.Print "Advertencia: No se pudo aplicar el estilo "
            Err.Clear
        End If
        On Error GoTo ErrorHandler
        sel.Font.Reset
        sel.ParagraphFormat.Reset

        
        For i = 0 To UBound(lines)
            lineText = lines(i)
            trimmedLine = Trim(lineText)
            trimmedLine = Replace(trimmedLine, vbLf & vbLf, vbLf)

            
            If Not inCodeBlock Then
                 
                Dim isBulleted As Boolean: isBulleted = (Left(trimmedLine, 1) = "*" Or Left(trimmedLine, 1) = "-" Or Left(trimmedLine, 1) = "+") And Len(trimmedLine) > 1
                Dim isNumbered As Boolean: isNumbered = False
                If IsNumeric(Left(trimmedLine, 1)) Then
                    listMarkerPos = InStr(trimmedLine, ". ")
                    If listMarkerPos > 1 And listMarkerPos <= Len(Left(trimmedLine, 1)) + 2 Then
                        isNumbered = True
                    End If
                End If

                If Len(trimmedLine) = 0 Or Not (isBulleted Or isNumbered) Then
                    If currentListType <> "" Then
    On Error Resume Next
    sel.Range.ListFormat.RemoveNumbers NumberType:=wdListNoNumbering
    If Err.Number <> 0 Then Debug.Print "Advertencia leve: No se pudo quitar formato de lista al final. Error: " & Err.Description: Err.Clear

    If Len(Trim(sel.Paragraphs(1).Range.text)) <= 1 Then
        sel.ParagraphFormat.Reset
        sel.Font.Reset
    End If
    On Error GoTo ErrorHandler
    currentListType = ""
    Debug.Print "Fin forzado de lista al final del procesamiento."
End If
                    lastLineWasList = False
                End If
            End If


            If Len(trimmedLine) = 0 And Not inCodeBlock Then
                
                
                Dim currentParaText As String
                On Error Resume Next
                currentParaText = sel.Paragraphs(1).Range.text
                If Err.Number <> 0 Then currentParaText = "Error": Err.Clear
                On Error GoTo ErrorHandler

                If Len(Trim(currentParaText)) > 1 Then
                   'sel.TypeParagraph
                   Debug.Print "Insertando p�rrafo vac�o."
                   
                   On Error Resume Next
                   If CustomStyle Then sel.Range.Style = WordDoc.Styles("Normal")
                   If Err.Number <> 0 Then Debug.Print "Adv: No se pudo aplicar estilo Normal a p�rrafo vac�o.": Err.Clear
                   On Error GoTo ErrorHandler
                   sel.Font.Reset
                   sel.ParagraphFormat.Reset
                Else
                   Debug.Print "P�rrafo ya vac�o, omitiendo inserci�n extra."
                   
                   
                   If lastLineWasList Then
                       On Error Resume Next
                       sel.Range.ListFormat.RemoveNumbers NumberType:=wdListNoNumbering
                       sel.ParagraphFormat.Reset
                       sel.Font.Reset
                       Err.Clear
                       On Error GoTo ErrorHandler
                       currentListType = ""
                       lastLineWasList = False
                   End If
                End If
                GoTo NextLine
            End If

            
            If trimmedLine = "```" Then
                inCodeBlock = Not inCodeBlock
                If inCodeBlock Then
                    
                    Dim isEmptyParaCodeStart As Boolean
                    On Error Resume Next
                    isEmptyParaCodeStart = (Len(Trim(sel.Paragraphs(1).Range.text)) <= 1)
                    If Err.Number <> 0 Then isEmptyParaCodeStart = False: Err.Clear
                    On Error GoTo ErrorHandler
                    'If Not isEmptyParaCodeStart Then sel.TypeParagraph

                    Set codeBlockStartRange = sel.Paragraphs(1).Range
                    
                    sel.Font.Name = "Courier New"
                    sel.Font.Size = 10
                    sel.Font.Bold = False
                    sel.Font.Italic = False
                    sel.ParagraphFormat.SpaceBefore = 6
                    sel.ParagraphFormat.SpaceAfter = 0
                    sel.ParagraphFormat.LeftIndent = WordApp.InchesToPoints(0.25)
                    sel.ParagraphFormat.RightIndent = WordApp.InchesToPoints(0.25)
                    Debug.Print "Inicio de bloque de c�digo detectado."
                Else
                    If Not codeBlockStartRange Is Nothing Then
                        Dim endParaRange As Object
                        Dim blockRange As Object
                        On Error Resume Next

                        
                        
                        
                        Set endParaRange = sel.Paragraphs(1).Previous.Range
                        If Err.Number <> 0 Or endParaRange Is Nothing Then
                           
                           
                           Set blockRange = codeBlockStartRange
                           Debug.Print "Advertencia: Bloque de c�digo corto o error al obtener p�rrafo anterior. Formateando solo el p�rrafo inicial."
                           Err.Clear
                        Else
                           
                           Set blockRange = WordDoc.Range(Start:=codeBlockStartRange.Start, End:=endParaRange.End)
                        End If
                        On Error GoTo ErrorHandler

                        If Not blockRange Is Nothing Then
                            
                            If blockRange.Start < blockRange.End Or blockRange.Characters.Count > 1 Then
                                On Error Resume Next
                                With blockRange.ParagraphFormat.Borders
                                     .Enable = True
                                     
                                     Dim borderType As Variant
                                     For Each borderType In Array(wdBorderTop, wdBorderLeft, wdBorderBottom, wdBorderRight)
                                         With .Item(CLng(borderType))
                                             .LineStyle = wdLineStyleSingle
                                             .LineWidth = wdLineWidth050pt
                                             .Color = wdColorGray25
                                         End With
                                     Next borderType
                                     
                                     .Item(wdBorderHorizontal).LineStyle = 0
                                     .Item(wdBorderVertical).LineStyle = 0
                                End With
                                blockRange.Shading.BackgroundPatternColor = RGB(248, 248, 248)
                                If Err.Number <> 0 Then Debug.Print "Adv: Error parcial al aplicar formato al bloque de c�digo. " & Err.Description: Err.Clear
                                Debug.Print "Bloque de c�digo formateado. Rango: " & blockRange.Start & "-" & blockRange.End
                            Else
                                Debug.Print "Advertencia: Rango de bloque de c�digo inv�lido o vac�o (" & blockRange.Start & "-" & blockRange.End & "). No se aplic� formato de borde/sombreado."
                            End If
                            Err.Clear
                            On Error GoTo ErrorHandler
                            Set blockRange = Nothing
                        End If

                        
                        sel.Collapse wdCollapseEnd
                        sel.TypeParagraph
                        sel.ParagraphFormat.Reset
                        sel.Font.Reset
                         
                        On Error Resume Next
                        sel.ParagraphFormat.Borders.Enable = False
                        sel.Shading.BackgroundPatternColor = wdColorAutomatic
                        Err.Clear
                        On Error GoTo ErrorHandler
                        Debug.Print "Fin de bloque de c�digo. Formato reseteado."
                    Else
                        Debug.Print "Advertencia: Se encontr� "
                        
                        sel.TypeParagraph
                    End If
                    Set codeBlockStartRange = Nothing
                End If
                lastLineWasList = False

            ElseIf inCodeBlock Then
                
                
                sel.TypeText text:=lineText
                sel.TypeParagraph
                
                sel.Font.Name = "Courier New"
                sel.Font.Size = 10
                sel.Font.Bold = False
                sel.Font.Italic = False
                sel.ParagraphFormat.LeftIndent = WordApp.InchesToPoints(0.25)
                sel.ParagraphFormat.RightIndent = WordApp.InchesToPoints(0.25)
                sel.ParagraphFormat.SpaceBefore = 0
                sel.ParagraphFormat.SpaceAfter = 0
                lastLineWasList = False

            
            ElseIf Left(trimmedLine, 1) = "#" Then
                hLevel = 0
                Do While Left(trimmedLine, hLevel + 1) Like String(hLevel + 1, "#") And hLevel < 6
                    hLevel = hLevel + 1
                Loop

                
                If hLevel > 0 And Mid(trimmedLine, hLevel + 1, 1) = " " Then
                    contentText = Trim(Mid(trimmedLine, hLevel + 2))
                    Debug.Print "Encabezado Nivel " & hLevel & " detectado: "

                    
                    Set paraRange = sel.Paragraphs(1).Range
                    If Len(Trim(paraRange.text)) > 1 Then sel.TypeParagraph

                    
                    sel.Font.Bold = True
                    
                    Select Case hLevel
                        Case 1: sel.Font.Size = 16
                        Case 2: sel.Font.Size = 14
                        Case 3: sel.Font.Size = 13
                        Case 4: sel.Font.Size = 12
                        Case Else: sel.Font.Size = 11
                    End Select
                    sel.ParagraphFormat.SpaceBefore = IIf(hLevel <= 2, 12, 6)
                    sel.ParagraphFormat.SpaceAfter = IIf(hLevel <= 3, 6, 4)

                    
                    
                    ProcessInlineFormatting sel, contentText, isHeader:=True

                    
                    sel.Collapse wdCollapseEnd
                    sel.TypeParagraph

                    
                    sel.Font.Reset
                    sel.ParagraphFormat.Reset
                    On Error Resume Next
                    If CustomStyle Then sel.Range.Style = WordDoc.Styles("Normal")
                    Err.Clear
                    On Error GoTo ErrorHandler

                    lastLineWasList = False
                Else
                    Debug.Print "Tratando l�nea que empieza con # pero no es encabezado como texto normal."
                     Set paraRange = sel.Paragraphs(1).Range
                     If Len(Trim(paraRange.text)) > 1 Then sel.TypeParagraph
                     sel.ParagraphFormat.Reset
                     sel.Font.Reset
                     ProcessInlineFormatting sel, trimmedLine
                     lastLineWasList = False
                End If

            
            ElseIf (Left(trimmedLine, 1) = "*" Or Left(trimmedLine, 1) = "-" Or Left(trimmedLine, 1) = "+") And Mid(trimmedLine, 2, 1) = " " Then
                contentText = Trim(Mid(trimmedLine, 3))
                Debug.Print "Elemento de lista con vi�eta detectado: "

                Set paraRange = sel.Paragraphs(1).Range
                'If Len(Trim(paraRange.text)) > 1 And Not lastLineWasList Then
               '     sel.TypeParagraph
                'End If

                If currentListType <> "bullet" Or Not lastLineWasList Then
                    On Error Resume Next
                    ' Aplica el formato de lista con vi�etas
                    Dim listGalleryBullet As Object
                    Set listGalleryBullet = WordApp.ListGalleries(wdBulletListNum)
                    sel.Range.ListFormat.ApplyListTemplateWithLevel ListTemplate:=listGalleryBullet.ListTemplates(1), ContinuePreviousList:=False, ApplyTo:=wdListApplyToSelection
                    If Err.Number <> 0 Then
                       Debug.Print "Error aplicando plantilla de vi�eta (" & Err.Description & "). Intentando ApplyBulletDefault."
                       Err.Clear
                       sel.Range.ListFormat.ApplyBulletDefault
                       If Err.Number <> 0 Then
                            Debug.Print "Error aplicando formato de vi�eta. Insertando texto con tabulaci�n."
                            Err.Clear
                            sel.TypeText vbTab & contentText ' Incluir el texto aqu�
                        Else
                           currentListType = "bullet"
                           Debug.Print "Aplicado formato de lista con vi�etas (ApplyBulletDefault)."
                           sel.TypeText " " & contentText ' A�adir el texto despu�s de la vi�eta autom�tica
                           sel.Collapse wdCollapseEnd
                           sel.TypeParagraph
                        End If
                    Else
                        currentListType = "bullet"
                        Debug.Print "Aplicado formato de lista con vi�etas (ApplyListTemplate)."
                        sel.TypeText " " & contentText ' A�adir el texto despu�s de la vi�eta autom�tica
                        sel.Collapse wdCollapseEnd
                        sel.TypeParagraph
                    End If
                    On Error GoTo ErrorHandler
                Else
                    ' Continuar la lista existente
                    sel.TypeText " " & contentText
                    sel.Collapse wdCollapseEnd
                    sel.TypeParagraph
                    Debug.Print "Continuando lista con vi�etas."
                End If
                lastLineWasList = True


            ElseIf isNumbered Then
                contentText = Trim(Mid(trimmedLine, listMarkerPos + 1))
                Debug.Print "Elemento de lista numerada detectado: "

                 Set paraRange = sel.Paragraphs(1).Range
                 'If Len(Trim(paraRange.text)) > 1 And Not lastLineWasList Then
                 '    sel.TypeParagraph
                ' End If

                 If currentListType <> "number" Or Not lastLineWasList Then
                    On Error Resume Next
                     ' Aplica el formato de lista numerada
                     Dim listGalleryNumber As Object
                     Set listGalleryNumber = WordApp.ListGalleries(wdNumberListNum)
                     sel.Range.ListFormat.ApplyListTemplateWithLevel ListTemplate:=listGalleryNumber.ListTemplates(1), ContinuePreviousList:=False, ApplyTo:=wdListApplyToSelection, DefaultListBehavior:=wdWord10ListBehavior
                    If Err.Number <> 0 Then
                         Debug.Print "Error aplicando plantilla numerada (" & Err.Description & "). Intentando ApplyNumberDefault."
                         Err.Clear
                         sel.Range.ListFormat.ApplyNumberDefault
                         If Err.Number <> 0 Then
                            Debug.Print "Error aplicando formato numerado. Insertando texto con n�mero y tab."
                            Err.Clear
                            sel.TypeText Trim(Left(trimmedLine, listMarkerPos)) & vbTab & contentText ' Incluir el texto
                            sel.Collapse wdCollapseEnd
                            sel.TypeParagraph
                         Else
                            currentListType = "number"
                            Debug.Print "Aplicado formato de lista numerada (ApplyNumberDefault)."
                            sel.TypeText " " & contentText ' A�adir el texto despu�s del n�mero autom�tico
                            sel.Collapse wdCollapseEnd
                            sel.TypeParagraph
                        End If
                    Else
                       currentListType = "number"
                       Debug.Print "Aplicado formato de lista numerada (ApplyListTemplate)."
                       sel.TypeText " " & contentText ' A�adir el texto despu�s del n�mero autom�tico
                       sel.Collapse wdCollapseEnd
                       sel.TypeParagraph
                    End If
                    On Error GoTo ErrorHandler
                 Else
                    ' Continuar la lista existente
                    sel.TypeText " " & contentText
                    sel.Collapse wdCollapseEnd
                    sel.TypeParagraph
                    Debug.Print "Continuando lista numerada."
                 End If

                lastLineWasList = True

            
            Else
                Debug.Print "Procesando como texto normal: "
                Set paraRange = sel.Paragraphs(1).Range
                
                 If Len(Trim(paraRange.text)) > 1 And (Not lastLineWasList Or i = 0) Then
                    'sel.TypeParagraph
                    sel.ParagraphFormat.Reset
                    sel.Font.Reset
                 End If
                 
                 If lastLineWasList Then
                    sel.ParagraphFormat.Reset
                    sel.Font.Reset
                 End If

                ProcessInlineFormatting sel, trimmedLine
                lastLineWasList = False
                currentListType = ""
            End If

NextLine:
        Next i
                
        
        If inCodeBlock Then
            Debug.Print "Advertencia: El texto Markdown termin� dentro de un bloque de c�digo sin cierre "
            
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
            
            sel.Collapse wdCollapseEnd
            sel.TypeParagraph
            sel.Font.Reset
            sel.ParagraphFormat.Reset
            On Error Resume Next
            sel.ParagraphFormat.Borders.Enable = False
            sel.Shading.BackgroundPatternColor = wdColorAutomatic
            Err.Clear
            On Error GoTo ErrorHandler
        End If

        
        If currentListType <> "" Then
            On Error Resume Next
             Set paraRange = sel.Paragraphs(1).Range
             If Len(Trim(paraRange.text)) <= 1 Then
                 paraRange.ListFormat.RemoveNumbers NumberType:=wdListNoNumbering
                 paraRange.ParagraphFormat.Reset
                 If Err.Number <> 0 Then Debug.Print "Advertencia: No se pudo quitar formato de lista final expl�citamente. Error: " & Err.Description: Err.Clear
             End If
            On Error GoTo ErrorHandler
        End If

        Debug.Print "Procesamiento de Markdown completado."

    Else
        MsgBox "Error: No se encontr� el placeholder "
    End If

    
    Set sel = Nothing
    Set fs = Nothing
    Set codeBlockStartRange = Nothing
    Set paraRange = Nothing

    Exit Sub

ErrorHandler:
    
    Dim errMsg As String
    errMsg = "Se produjo un error inesperado en " & _
             "Error N�mero: " & Err.Number & vbCrLf & _
             "Descripci�n: " & Err.Description & vbCrLf & _
             "Fuente: " & Err.Source & vbCrLf & vbCrLf
    
    On Error Resume Next
    errMsg = errMsg & "�ltima l�nea procesada (aprox.): " & i
    If i >= 0 And i <= UBound(lines) Then
       errMsg = errMsg & " -> "
    End If
    If Err.Number <> 0 Then
        errMsg = errMsg & " (No se pudo determinar el contenido de la l�nea)."
        Err.Clear
    End If
    On Error GoTo 0

    MsgBox errMsg, vbCritical, "Error en Ejecuci�n VBA"

    
    On Error Resume Next
    Set sel = Nothing
    Set fs = Nothing
    Set codeBlockStartRange = Nothing
    Set paraRange = Nothing
    On Error GoTo 0

End Sub



Private Sub ProcessInlineFormatting(sel As Object, text As String, Optional isListItem As Boolean = False, Optional isHeader As Boolean = False)
    
    Const BOLD_MARKER As String = "**"
    Const ITALIC_MARKER As String = "*"
    Const CODE_MARKER As String = "`"
    Const ESCAPE_CHAR As String = "\"
    Const CODE_FONT_NAME As String = "Courier New"
    Const CODE_FONT_SIZE As Long = 10
    Const WD_COLLAPSE_END As Long = 0
    Dim wdColorAutomaticInline As Long: wdColorAutomaticInline = -16777216

    
    Dim currPos As Long
    Dim char As String, nextChar As String, prevChar As String
    Dim textToInsert As String
    Dim initialFont As Object
    Dim initialBold As Boolean, initialItalic As Boolean
    Dim isCurrentlyCode As Boolean
    Dim tempRange As Object

    On Error GoTo InlineErrorHandler

    
    
    
    Set tempRange = sel.Range
    Set initialFont = tempRange.Font.Duplicate
    Set tempRange = Nothing
    isCurrentlyCode = False

    
    currPos = 1
    textToInsert = ""

    Do While currPos <= Len(text)
        char = Mid(text, currPos, 1)
        
        If currPos < Len(text) Then nextChar = Mid(text, currPos + 1, 1) Else nextChar = ""
        If currPos > 1 Then prevChar = Mid(text, currPos - 1, 1) Else prevChar = ""

        
        If char = ESCAPE_CHAR Then
             
            If nextChar = Left(BOLD_MARKER, 1) Or nextChar = ITALIC_MARKER Or nextChar = CODE_MARKER Or nextChar = ESCAPE_CHAR Then
                textToInsert = textToInsert & nextChar
                currPos = currPos + 2
                GoTo ContinueLoopInline
            Else
                textToInsert = textToInsert & char
                
                currPos = currPos + 1
                GoTo ContinueLoopInline
            End If
        End If

        
        Dim markerFound As Boolean: markerFound = False

        
        If char = ITALIC_MARKER And nextChar = ITALIC_MARKER Then
            If Len(textToInsert) > 0 Then sel.TypeText textToInsert: textToInsert = ""
            sel.Font.Bold = Not sel.Font.Bold
            currPos = currPos + 2
            markerFound = True
        
        ElseIf char = ITALIC_MARKER Then
            
            If prevChar <> ITALIC_MARKER Then
                 If Len(textToInsert) > 0 Then sel.TypeText textToInsert: textToInsert = ""
                 sel.Font.Italic = Not sel.Font.Italic
                 currPos = currPos + 1
                 markerFound = True
            Else
                 currPos = currPos + 1
                 markerFound = True
            End If
        
        ElseIf char = CODE_MARKER Then
            If Len(textToInsert) > 0 Then sel.TypeText textToInsert: textToInsert = ""
            isCurrentlyCode = Not isCurrentlyCode
            If isCurrentlyCode Then
                
                initialBold = sel.Font.Bold
                initialItalic = sel.Font.Italic
                
                sel.Font.Name = CODE_FONT_NAME
                sel.Font.Size = CODE_FONT_SIZE
                sel.Font.Bold = False
                sel.Font.Italic = False
                 
                 sel.Shading.BackgroundPatternColor = RGB(240, 240, 240)
            Else
                
                 If Not initialFont Is Nothing Then
                     sel.Font.Name = initialFont.Name
                     sel.Font.Size = initialFont.Size
                     sel.Shading.BackgroundPatternColor = wdColorAutomaticInline
                     
                     sel.Font.Bold = initialBold
                     sel.Font.Italic = initialItalic
                 Else
                     sel.Font.Reset
                     sel.Shading.BackgroundPatternColor = wdColorAutomaticInline
                 End If
            End If
            currPos = currPos + 1
            markerFound = True
        End If

        
        If Not markerFound Then
            textToInsert = textToInsert & char
            currPos = currPos + 1
        End If

ContinueLoopInline:
    Loop

    
    If Len(textToInsert) > 0 Then sel.TypeText textToInsert

    
    If Not isListItem And Not isHeader Then
        
        sel.Collapse WD_COLLAPSE_END
        sel.TypeParagraph
    End If
    

    Set initialFont = Nothing
    Exit Sub

InlineErrorHandler:
    MsgBox "Error en " & _
           "Error: " & Err.Number & " - " & Err.Description & vbCrLf & _
           "Procesando texto (inicio):"
    
    On Error Resume Next
    sel.TypeText text
    If Not isListItem And Not isHeader Then sel.TypeParagraph
    Err.Clear
    On Error GoTo 0
    Set initialFont = Nothing
End Sub





Sub FusionarDocumentosInsertando(WordApp As Object, documentsList As Variant, finalDocumentPath As String)
    Dim baseDoc     As Object
    Dim sFile       As String
    Dim oRng        As Object
    Dim i           As Integer
    
    On Error GoTo err_Handler
    
    
    Set baseDoc = WordApp.Documents.Add
    
    
    For i = LBound(documentsList) To UBound(documentsList)
        sFile = documentsList(i)
        
        
        Set oRng = baseDoc.Range
        oRng.Collapse 0
        oRng.InsertFile sFile, , , , True
        
        
        If i < UBound(documentsList) Then
            Set oRng = baseDoc.Range
            oRng.Collapse 0
            oRng.InsertBreak Type:=6
        End If
    Next i
        
    
    baseDoc.SaveAs finalDocumentPath
    
    
    baseDoc.Close
    
    
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
    
    
    If fs Is Nothing Then
        Debug.Print "FSO inv�lido en IsAbsolutePath"
        Exit Function
    End If

    
    Dim lowerPath As String: lowerPath = LCase(Trim(path))

    
    If Left(lowerPath, 7) = "http://" Or Left(lowerPath, 8) = "https://" Then
        IsAbsolutePath = True
        GoTo CleanExitAbsPath
    End If

    
    Dim driveName As String
    driveName = fs.GetDriveName(path)

    
    If Err.Number = 0 Then
        
        If Left(driveName, 2) = "\\" Then
            IsAbsolutePath = True
        
        ElseIf Len(driveName) = 2 And Right(driveName, 1) = ":" Then
            If Asc(LCase(Left(driveName, 1))) >= 97 And Asc(LCase(Left(driveName, 1))) <= 122 Then
                IsAbsolutePath = True
            End If
        End If
    End If
    Err.Clear

CleanExitAbsPath:
    
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
    
    ' Reemplazar saltos de l�nea y salto de carro con s�mbolos visibles
    formattedText = Replace(formattedText, vbCrLf, "[CRLF]")
    formattedText = Replace(formattedText, vbCr, "[CR]")
    formattedText = Replace(formattedText, vbLf, "[LF]")
    
    ' Reemplazar otros caracteres especiales con su representaci�n correspondiente
    formattedText = Replace(formattedText, Chr(9), "[TAB]") ' TAB
    formattedText = Replace(formattedText, Chr(0), "[NULL]") ' NULL
    
    ' Imprimir el texto con los saltos de l�nea visibles
    Debug.Print "Texto formateado con saltos visibles:" & vbCrLf & formattedText

End Sub

Public Sub SustituirTextoMarkdownPorImagenes(WordApp As Object, WordDoc As Object, basePath As String)
    Dim docRange As Object
    Dim searchRange As Object
    Dim tagRange As Object
    Dim expandedTagRange As Object
    Dim fs As Object
    Dim imgAltText As String
    Dim imgPath As String
    Dim fullImgPath As String
    Dim altEndMarkerPos As Long
    Dim pathEndMarkerPos As Long
    Dim inlineShape As Object
    Dim originalTagText As String
    Dim fileExists As Boolean
    Dim continueSearching As Boolean
    Dim pathSep As String
    Dim charCodeBefore As Long
    Dim charCodeAfter As Long
    Dim currentSearchStart As Long
    Dim tempSearchRange As Object
    Dim coreTagText As String
    Dim relAltEndPos As Long, relPathStartPos As Long
    Dim insertionPointRange As Object
    Dim rngAfterPic As Object
    Dim addPictureErrNum As Long
    Dim pathIsAbsolute As Boolean
    Dim rngBeforePic As Object ' Declaraci�n para el p�rrafo antes de la imagen
    Dim rngCaption As Object    ' Declaraci�n para el rango del caption
    Dim rngAfterCaption As Object ' Declaraci�n para el p�rrafo despu�s del caption

    Const CR As Long = 13
    Const LF As Long = 10

    Const START_MARKER As String = "!["
    Const ALT_END_MARKER As String = "]("
    Const PATH_END_MARKER As String = ")"

    Const wdFindStop As Long = 0
    Const wdCollapseStart As Long = 1
    Const wdCollapseEnd As Long = 0
    Const wdAlignParagraphCenter As Long = 1
    Const wdAlignParagraphJustify As Long = 3 ' Constante para justificar
    Const wdParagraph As Long = 4

    On Error GoTo GlobalErrorHandler

    If WordApp Is Nothing Or WordDoc Is Nothing Then
        MsgBox "Error: La aplicaci�n Word o el Documento no son v�lidos.", vbCritical, "Error de Entrada"
        Exit Sub
    End If

    On Error Resume Next
    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs Is Nothing Or Err.Number <> 0 Then
        On Error GoTo 0
        MsgBox "Error cr�tico: No se pudo crear el FileSystemObject.", vbCritical, "Error FSO"
        Exit Sub
    End If
    On Error GoTo GlobalErrorHandler

    pathSep = WordApp.PathSeparator
    basePath = Trim(basePath)
    If Len(basePath) > 0 Then
        ' Normalizar la basePath eliminando barras diagonales al final
        Do While Right(basePath, 1) = "\" Or Right(basePath, 1) = "/"
            If Len(basePath) = 1 Then
                basePath = ""
                Exit Do
            End If
            basePath = Left(basePath, Len(basePath) - 1)
        Loop
        ' A�adir separador de ruta si basePath no est� vac�o
        If Len(basePath) > 0 Then
            basePath = basePath & pathSep
        End If
        Debug.Print "BasePath normalizado: """ & basePath & """"
    End If

    currentSearchStart = 0
    continueSearching = True
    WordApp.ScreenUpdating = False

    Do While continueSearching
        Set searchRange = WordDoc.Range(Start:=currentSearchStart, End:=WordDoc.content.End)

        searchRange.Find.ClearFormatting
        searchRange.Find.text = START_MARKER
        searchRange.Find.Forward = True
        searchRange.Find.Wrap = wdFindStop
        searchRange.Find.MatchCase = False

        If searchRange.Find.Execute Then
            Dim foundRange As Object
            Set foundRange = searchRange.Duplicate

            currentSearchStart = foundRange.End

            Set tempSearchRange = WordDoc.Range(Start:=foundRange.End, End:=WordDoc.content.End)

            altEndMarkerPos = InStr(1, tempSearchRange.text, ALT_END_MARKER, vbTextCompare)

            If altEndMarkerPos > 0 Then
                pathEndMarkerPos = InStr(altEndMarkerPos + Len(ALT_END_MARKER), tempSearchRange.text, PATH_END_MARKER, vbTextCompare)

                If pathEndMarkerPos > 0 Then
                    Dim tagStartDocPos As Long, tagEndDocPos As Long
                    tagStartDocPos = foundRange.Start

                    tagEndDocPos = tempSearchRange.Start + pathEndMarkerPos - 1 + Len(PATH_END_MARKER)

                    Set tagRange = WordDoc.Range(Start:=tagStartDocPos, End:=tagEndDocPos)
                    coreTagText = tagRange.text

                    relAltEndPos = InStr(1, coreTagText, ALT_END_MARKER, vbTextCompare)
                    If relAltEndPos = 0 Then
                        Debug.Print "Error interno: No se encontr� ALT_END_MARKER en coreTagText. Saltando."
                        GoTo ContinueNextFind
                    End If

                    relPathStartPos = relAltEndPos + Len(ALT_END_MARKER)
                    imgAltText = Trim(Mid(coreTagText, Len(START_MARKER) + 1, relAltEndPos - Len(START_MARKER) - 1))
                    imgPath = Trim(Mid(coreTagText, relPathStartPos, pathEndMarkerPos - relPathStartPos))

                    Dim relPathEndPos As Long
                    relPathEndPos = InStr(relPathStartPos, coreTagText, PATH_END_MARKER, vbTextCompare)
                    If relPathEndPos > 0 Then
                        imgPath = Trim(Mid(coreTagText, relPathStartPos, relPathEndPos - relPathStartPos))
                    Else
                        Debug.Print "Error interno: No se encontr� PATH_END_MARKER en coreTagText. Saltando."
                        GoTo ContinueNextFind
                    End If

                    Debug.Print "Tag encontrado: """ & Replace(Replace(coreTagText, Chr(CR), "[CR]"), Chr(LF), "[LF]") & """"
                    Debug.Print "Texto Alt extra�do: """ & imgAltText & """"
                    Debug.Print "Ruta extra�da (cruda): """ & imgPath & """"

                    If Len(imgPath) = 0 Then
                        Debug.Print "Advertencia: Ruta de imagen vac�a. Saltando tag."
                        currentSearchStart = tagRange.End
                        GoTo ContinueNextFind
                    End If

                    Set expandedTagRange = tagRange.Duplicate
                    originalTagText = expandedTagRange.text
                    Debug.Print "Rango tag inicial: " & expandedTagRange.Start & "-" & expandedTagRange.End

                    ' Expandir el rango para incluir la l�nea anterior si es un CR
                    If expandedTagRange.Start > 0 Then
                        On Error Resume Next
                        Dim rngBefore As Object
                        Set rngBefore = WordDoc.Range(expandedTagRange.Start - 1, expandedTagRange.Start)
                        If Err.Number = 0 Then
                            charCodeBefore = AscW(rngBefore.Characters(1).text)
                            If charCodeBefore = CR Then
                                expandedTagRange.Start = expandedTagRange.Start - 1
                                Debug.Print "CR encontrado antes. Rango expandido a: " & expandedTagRange.Start
                            End If
                        End If
                        Set rngBefore = Nothing
                        Err.Clear
                        On Error GoTo GlobalErrorHandler
                    End If

                    ' Expandir el rango para incluir la l�nea posterior si es un CR
                    If expandedTagRange.End < WordDoc.content.End Then
                        On Error Resume Next
                        Dim rngAfter As Object
                        Set rngAfter = WordDoc.Range(expandedTagRange.End, expandedTagRange.End + 1)
                        If Err.Number = 0 Then
                            charCodeAfter = AscW(rngAfter.Characters(1).text)
                            If charCodeAfter = CR Then
                                expandedTagRange.End = expandedTagRange.End + 1
                                Debug.Print "CR encontrado despu�s. Rango expandido a: " & expandedTagRange.End
                            End If
                        End If
                        Set rngAfter = Nothing
                        Err.Clear
                        On Error GoTo GlobalErrorHandler
                    End If
                    Debug.Print "Texto a reemplazar (expandido): """ & Replace(Replace(expandedTagRange.text, Chr(CR), "[CR]"), Chr(LF), "[LF]") & """"

                    fullImgPath = ""
                    fileExists = False

                    imgPath = Replace(imgPath, "/", pathSep)
                    imgPath = Replace(imgPath, "\", pathSep)

                    On Error Resume Next
                    pathIsAbsolute = fs.IsAbsolutePath(imgPath)
                    If Err.Number <> 0 Then
                        Err.Clear
                        Debug.Print "Advertencia: fs.IsAbsolutePath fall�. Usando chequeo manual."
                        pathIsAbsolute = (InStr(imgPath, ":\") > 0 Or Left(imgPath, 2) = "\\")
                    End If
                    On Error GoTo GlobalErrorHandler

                    If pathIsAbsolute Then
                        fullImgPath = imgPath
                        On Error Resume Next
                        fileExists = fs.fileExists(fullImgPath)
                        If Err.Number <> 0 Then
                            Debug.Print "Error FSO chequeando existencia de: " & fullImgPath
                            fileExists = False
                            Err.Clear
                        End If
                        On Error GoTo GlobalErrorHandler
                    ElseIf Left(LCase(imgPath), 4) = "http" Then
                        fullImgPath = imgPath
                        fileExists = True
                        Debug.Print "Ruta es URL: """ & fullImgPath & """"
                    ElseIf Len(basePath) > 0 Then
                        On Error Resume Next
                        fullImgPath = fs.BuildPath(basePath, imgPath)
                        If Err.Number = 0 Then
                            fileExists = fs.fileExists(fullImgPath)
                            If Err.Number <> 0 Then
                                Debug.Print "Error FSO chequeando existencia de ruta construida: " & fullImgPath
                                fileExists = False
                                Err.Clear
                            End If
                        Else
                            Debug.Print "Error FSO en BuildPath con: """ & basePath & """ y """ & imgPath & """"
                            fullImgPath = basePath & imgPath
                            fileExists = False
                            Err.Clear
                        End If
                        On Error GoTo GlobalErrorHandler
                        Debug.Print "Ruta relativa construida: """ & fullImgPath & """"
                    Else
                        Debug.Print "Advertencia: Ruta relativa (""" & imgPath & """) encontrada pero no se proporcion� BasePath."
                        fullImgPath = imgPath
                        fileExists = False
                    End If

                    If fileExists And Len(fullImgPath) > 0 Then
                        Debug.Print "Archivo/URL parece existir: """ & fullImgPath & """. Intentando insertar..."
                        
                             ' Primero, borrar el texto del tag antes de hacer cualquier inserci�n
    expandedTagRange.text = ""
    
    ' Crear rango colapsado al inicio del tag original (ya vac�o)
    Set insertionPointRange = expandedTagRange.Duplicate
    insertionPointRange.Collapse Direction:=wdCollapseStart

    ' Insertar un p�rrafo vac�o antes
    insertionPointRange.InsertParagraphBefore

    ' Mover el rango al p�rrafo insertado para insertar imagen justo despu�s
    Set rngBeforePic = insertionPointRange.Paragraphs(1).Range
    rngBeforePic.Collapse Direction:=wdCollapseEnd


                        Set inlineShape = Nothing
                        addPictureErrNum = 0
                        On Error Resume Next
                        Set inlineShape = rngBeforePic.InlineShapes.AddPicture( _
                            fileName:=fullImgPath, LinkToFile:=False, SaveWithDocument:=True)
                        addPictureErrNum = Err.Number
                        On Error GoTo GlobalErrorHandler

                        If addPictureErrNum = 0 And Not inlineShape Is Nothing Then
                            Debug.Print "Imagen insertada con �xito."
                            inlineShape.AlternativeText = imgAltText

                            ' Centrar la imagen
                            inlineShape.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter

                            Set rngCaption = inlineShape.Range
                            rngCaption.Collapse Direction:=wdCollapseEnd

                            ' L�nea vac�a DESPU�S de la imagen (antes del caption)
                            rngCaption.InsertParagraphAfter
                            
                            
                                 ' Mover al rango del caption
                            Set rngCaption = rngCaption.Next(wdParagraph)
                            rngCaption.text = imgAltText
                            rngCaption.ParagraphFormat.Alignment = wdAlignParagraphCenter
                            rngCaption.Font.Italic = True
                            rngCaption.InsertParagraphAfter
                            'rngCaption.Font.Size = 9
                            
                              ' L�nea vac�a DESPU�S del caption
                            Set rngAfterCaption = rngCaption.Duplicate
                            rngAfterCaption.Collapse Direction:=wdCollapseEnd
                            rngAfterCaption.ParagraphFormat.Alignment = wdAlignParagraphJustify
                            rngAfterCaption.Font.Italic = False

                            currentSearchStart = rngAfterCaption.End
                        Else
                            Debug.Print "Error al insertar la imagen (Error #" & addPictureErrNum & ")."
                        End If
                    Else
                        Debug.Print "Archivo/URL no existe o est� vac�o: """ & fullImgPath & """."
                    End If
                Else
                    Debug.Print "Tag incompleto: Se encontr� '![' pero no se encontr� '](' o ')'. Saltando."
                    currentSearchStart = tempSearchRange.Start + altEndMarkerPos + Len(ALT_END_MARKER) - 1
                End If
            Else
                Debug.Print "Tag incompleto: Se encontr� '![' pero no se encontr� ']('. Continuando b�squeda."
                currentSearchStart = foundRange.End + Len(START_MARKER)
            End If

            Set tempSearchRange = Nothing
            Set foundRange = Nothing

        Else
            continueSearching = False
            Debug.Print "No se encontraron m�s marcadores '!['."
        End If

ContinueNextFind:
        DoEvents
    Loop

    GoTo CleanExit

GlobalErrorHandler:
    MsgBox "Se produjo un error inesperado en SustituirTextoMarkdownPorImagenes." & vbCrLf & _
            "Error " & Err.Number & ": " & Err.Description & vbCrLf & _
            "Fuente: " & Err.Source & vbCrLf & vbCrLf & _
            "�ltima ruta procesada (aprox): """ & fullImgPath & """", vbCritical, "Error Global en Ejecuci�n"

    If Not WordApp Is Nothing Then
        On Error Resume Next
        WordApp.ScreenUpdating = True
        On Error GoTo 0
    End If
    GoTo ForcedCleanup

CleanExit:
    Debug.Print "--- Fin de SustituirTextoMarkdownPorImagenes (Normal) ---"

ForcedCleanup:
    On Error Resume Next
    If Not fs Is Nothing Then Set fs = Nothing
    Set docRange = Nothing
    Set searchRange = Nothing
    Set tagRange = Nothing
    Set expandedTagRange = Nothing
    Set inlineShape = Nothing
    Set tempSearchRange = Nothing
    Set insertionPointRange = Nothing
    Set rngAfterPic = Nothing
    Set rngBeforePic = Nothing ' Liberar objeto
    Set rngCaption = Nothing    ' Liberar objeto
    Set rngAfterCaption = Nothing ' Liberar objeto
    If Not WordApp Is Nothing Then WordApp.ScreenUpdating = True
    On Error GoTo 0
End Sub




Sub AplicarFormatoCeldaEnTablaWord(WordDoc As Object, tablenum As Integer, row As Integer, col As Integer, styleName As String)
  
  Dim cellRange As Object
  Set cellRange = WordDoc.Tables(tablenum).cell(row, col).Range
  cellRange.Style = styleName
  
End Sub


Private Sub EliminarLineasVaciasEnCeldaTablaWord(WordDoc As Object, tablenum As Integer, row As Integer, col As Integer)
    With WordDoc.Tables(tablenum).cell(row, col).Range.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .text = "^p^p"
        .Replacement.text = "^p"
        .Forward = True
        .Wrap = False
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        
        ' Repetir hasta que ya no se encuentren m�s saltos m�ltiples
        Do While .Execute(Replace:=wdReplaceAll)
        Loop
    End With

    '  MsgBox WordDoc.Tables(tablenum).cell(row, col).Range.text
End Sub

Sub AjustarMarcadorCeldaEnTablaWord(ByRef WordApp As Object, ByRef WordDoc As Object, _
                            ByVal tablenum As Integer, ByVal row As Integer, ByVal col As Integer)
    Dim cell As Object
    Dim rng As Object
    Dim paraCount As Long
    Dim firstPara As Object
    Dim lastPara As Object
    

    ' A�ade estas constantes al inicio de tu m�dulo (necesarias para Excel VBA)
    Const wdLine = 5
    Const wdWord = 2

    ' Obtener la celda
    Set cell = WordDoc.Tables(tablenum).cell(row, col)

    ' Obtener el rango de la celda
    Set rng = cell.Range

    ' Contar los p�rrafos dentro de la celda
    paraCount = rng.Paragraphs.Count

    ' Si hay p�rrafos, mover el cursor al inicio del primer p�rrafo
    If paraCount > 0 Then
        ' Obtener el primer p�rrafo
        Set firstPara = rng.Paragraphs(1)
        Set lastPara = rng.Paragraphs(paraCount)
     
        ' Mover el cursor al inicio del primer p�rrafo
        rng.SetRange Start:=firstPara.Range.Start, End:=firstPara.Range.Start

        ' Seleccionar el rango (coloca el cursor all�)
        rng.SetRange Start:=lastPara.Range.End - 2, End:=lastPara.Range.End

        ' Mover el cursor hacia abajo una l�nea
        'WordApp.Selection.MoveDown Unit:=wdLine, Count:=1
        'MsgBox 1
        WordApp.Selection.TypeBackspace
    End If
End Sub


Function EliminarLineasVaciasdeString(explicacionTecnicaValue As String) As String
    
    
    Dim result As String
    Dim lines() As String
    Dim i As Integer
    Dim cleanText As String
     
    
    result = Replace(explicacionTecnicaValue, vbLf & vbLf, vbLf)
    result = Replace(result, vbCr & vbCr, vbCr)
    result = Replace(result, vbCrLf & vbCrLf, vbCrLf)
    
    
    EliminarLineasVaciasdeString = result
End Function


