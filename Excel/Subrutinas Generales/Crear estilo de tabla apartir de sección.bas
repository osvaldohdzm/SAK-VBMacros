Attribute VB_Name = "Module2"
Sub CrearNuevoEstiloDesdeTabla()
    Dim selectedRange As Range
    Dim tableName As String
    Dim newStyleName As String
    Dim newStyle As Style
    
    ' Paso 1: Seleccionar una tabla en Excel
    On Error Resume Next
    Set selectedRange = Application.InputBox("Seleccione una tabla en Excel", Type:=8)
    On Error GoTo 0
    
    If selectedRange Is Nothing Then
        MsgBox "No se seleccionó ninguna tabla. La macro se cancelará."
        Exit Sub
    End If
    
    ' Obtener el nombre de la tabla seleccionada
    tableName = selectedRange.Worksheet.ListObjects(1).Name
    
    ' Mostrar el nombre de la tabla en un MsgBox
    MsgBox "Nombre de la tabla seleccionada: " & tableName
    
    ' Paso 2: Obtener todas las propiedades de estilo de la tabla seleccionada
    Dim fontColor As Long
    Dim fontBold As Boolean
    Dim fontName As String
    Dim fontItalic As Boolean
    Dim fontSize As Double
    Dim fontStrikethrough As Boolean
    Dim fontSubscript As Boolean
    Dim fontSuperscript As Boolean
    
    Dim interiorColor As Long
    Dim interiorPattern As XlPattern
    Dim interiorPatternColor As Long
    
    Dim borderLineStyle As XlLineStyle
    Dim borderColor As Long
    Dim borderWidth As XlBorderWeight
    
    Dim addIndent As Boolean
    Dim formulaHidden As Boolean
    Dim horizontalAlignment As XlHAlign
    Dim indentLevel As Integer
    Dim numberFormat As String
    Dim numberFormatLocal As String
    Dim orientation As XlOrientation
    Dim shrinkToFit As Boolean
    Dim verticalAlignment As XlVAlign
    Dim wrapText As Boolean
    
    With selectedRange
        ' Propiedades de fuente
        fontColor = .Font.Color
        fontBold = .Font.Bold
        fontName = .Font.Name
        fontItalic = .Font.Italic
        fontSize = .Font.Size
        fontStrikethrough = .Font.Strikethrough
        fontSubscript = .Font.Subscript
        fontSuperscript = .Font.Superscript
        
        ' Propiedades de interior
        interiorColor = .Interior.Color
        interiorPattern = .Interior.Pattern
        interiorPatternColor = .Interior.PatternColor
        
        ' Propiedades de bordes (solo el primer borde en este ejemplo)
        If .Borders.Count > 0 Then
            borderLineStyle = .Borders(1).LineStyle
            borderColor = .Borders(1).Color
            borderWidth = .Borders(1).Weight
        End If
        
        ' Otras propiedades de estilo
        addIndent = .addIndent
        formulaHidden = .formulaHidden
        horizontalAlignment = .horizontalAlignment
        indentLevel = .indentLevel
        numberFormat = .numberFormat
        numberFormatLocal = .numberFormatLocal
        orientation = .orientation
        shrinkToFit = .shrinkToFit
        verticalAlignment = .verticalAlignment
        wrapText = .wrapText
    End With
    
    ' Paso 3: Crear un nuevo estilo con las propiedades obtenidas
    newStyleName = "Nuevo estilo desde tabla"
    Set newStyle = CreateNewStyle(newStyleName, fontColor, fontBold, fontName, fontItalic, _
                                  fontSize, fontStrikethrough, fontSubscript, fontSuperscript, _
                                  interiorColor, interiorPattern, interiorPatternColor, _
                                  borderLineStyle, borderColor, borderWidth, _
                                  addIndent, formulaHidden, horizontalAlignment, _
                                  indentLevel, numberFormat, numberFormatLocal, _
                                  orientation, shrinkToFit, verticalAlignment, wrapText)
    
    ' Informar al usuario que se ha creado el nuevo estilo
    MsgBox "Se ha creado el nuevo estilo '" & newStyleName & "' basado en la tabla seleccionada.", vbInformation
End Sub

Function CreateNewStyle(styleName As String, _
                        fontColor As Long, fontBold As Boolean, fontName As String, _
                        fontItalic As Boolean, fontSize As Double, fontStrikethrough As Boolean, _
                        fontSubscript As Boolean, fontSuperscript As Boolean, _
                        interiorColor As Long, interiorPattern As XlPattern, _
                        interiorPatternColor As Long, _
                        borderLineStyle As XlLineStyle, borderColor As Long, borderWidth As XlBorderWeight, _
                        addIndent As Boolean, formulaHidden As Boolean, horizontalAlignment As XlHAlign, _
                        indentLevel As Integer, numberFormat As String, numberFormatLocal As String, _
                        orientation As XlOrientation, shrinkToFit As Boolean, _
                        verticalAlignment As XlVAlign, wrapText As Boolean) As Style
    
    Dim newStyle As Style
    Dim newFont As Font
    Dim newInterior As Interior
    Dim newBorders As Borders
    
    ' Intentar obtener el estilo existente o crear uno nuevo si no existe
    On Error Resume Next
    Set newStyle = ThisWorkbook.Styles(styleName)
    On Error GoTo 0
    
    If newStyle Is Nothing Then
        Set newStyle = ThisWorkbook.Styles.Add(styleName)
    End If
    
    ' Configurar propiedades de estilo
    With newStyle
        ' Configurar propiedades de fuente
        Set newFont = .Font
        With newFont
            .Color = fontColor
            .Bold = fontBold
            .Name = fontName
            .Italic = fontItalic
            .Size = fontSize
            .Strikethrough = fontStrikethrough
            .Subscript = fontSubscript
            .Superscript = fontSuperscript
        End With
        
        ' Configurar propiedades de interior
        Set newInterior = .Interior
        With newInterior
            .Color = interiorColor
            .Pattern = interiorPattern
            .PatternColor = interiorPatternColor
        End With
        
        ' Configurar propiedades de bordes
        Set newBorders = .Borders
        With newBorders
            .LineStyle = borderLineStyle
            .Color = borderColor
            .Weight = borderWidth
        End With
        
        ' Configurar otras propiedades de estilo
        .addIndent = addIndent
        .formulaHidden = formulaHidden
        .horizontalAlignment = horizontalAlignment
        .indentLevel = indentLevel
        .numberFormat = numberFormat
        .numberFormatLocal = numberFormatLocal
        .orientation = orientation
        .shrinkToFit = shrinkToFit
        .verticalAlignment = verticalAlignment
        .wrapText = wrapText
    End With
    
    ' Devolver el estilo creado o modificado
    Set CreateNewStyle = newStyle
End Function


