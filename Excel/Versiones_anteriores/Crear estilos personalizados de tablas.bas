Attribute VB_Name = "Module1"
Sub CrearEstilosDeTablaConColores()
    ' Crear estilo con armonía azul
    CrearEstiloTablaConColor "Azul", RGB(0, 112, 192), RGB(0, 32, 96)
    
    ' Crear estilo con armonía roja
    CrearEstiloTablaConColor "Rojo", RGB(192, 0, 0), RGB(96, 0, 0)
    
    ' Crear estilo con armonía verde
    CrearEstiloTablaConColor "Verde", RGB(0, 176, 80), RGB(0, 88, 40)
    
    ' Crear estilo con armonía amarilla
    CrearEstiloTablaConColor "Amarillo", RGB(255, 192, 0), RGB(128, 96, 0)
    
    ' Crear estilo con armonía amarilla
    CrearEstiloTablaConColor "Tabla INAI 1", RGB(0, 32, 96), RGB(128, 96, 0)

    
    MsgBox "Se han creado los estilos de tabla con diferentes colores.", vbInformation
End Sub

Sub CrearEstiloTablaConColor(ByVal nombreEstilo As String, ByVal colorPrimario As Long, ByVal colorSecundario As Long)
    Dim newStyleName As String
    Dim newStyle As TableStyle
    
    ' Crear un nuevo estilo de tabla duplicando uno existente (en este caso TableStyleMedium2)
    ActiveWorkbook.TableStyles("TableStyleMedium2").Duplicate ("TableStyleMedium2_" & nombreEstilo)
    
    ' Obtener el nuevo estilo creado
    Set newStyle = ActiveWorkbook.TableStyles("TableStyleMedium2_" & nombreEstilo)
    
    ' Configurar el nuevo estilo con la armonía de colores
    With newStyle
        .ShowAsAvailablePivotTableStyle = False
        .ShowAsAvailableTableStyle = True
        .ShowAsAvailableSlicerStyle = False
        .ShowAsAvailableTimelineStyle = False
        
        ' Configurar elementos de estilo para toda la tabla
        With .TableStyleElements(xlWholeTable).Font
            .TintAndShade = 0
            .ThemeColor = xlThemeColorLight1
        End With
        
        With .TableStyleElements(xlWholeTable).Interior
            .Pattern = xlNone
            .TintAndShade = 0
        End With
        
        With .TableStyleElements(xlWholeTable).Borders
            .LineStyle = xlNone
        End With
        
        ' Configurar bordes específicos
        With .TableStyleElements(xlWholeTable).Borders(xlEdgeTop)
            .ThemeColor = xlThemeColorAccent1
            .TintAndShade = 0.399975585192419
            .Weight = xlThin
            .LineStyle = xlNone
        End With
        
        With .TableStyleElements(xlWholeTable).Borders(xlEdgeBottom)
            .ThemeColor = xlThemeColorAccent1
            .TintAndShade = 0.399975585192419
            .Weight = xlThin
            .LineStyle = xlNone
        End With
        
        With .TableStyleElements(xlWholeTable).Borders(xlEdgeLeft)
            .ThemeColor = xlThemeColorAccent1
            .TintAndShade = 0.399975585192419
            .Weight = xlThin
            .LineStyle = xlNone
        End With
        
        With .TableStyleElements(xlWholeTable).Borders(xlEdgeRight)
            .ThemeColor = xlThemeColorAccent1
            .TintAndShade = 0.399975585192419
            .Weight = xlThin
            .LineStyle = xlNone
        End With
        
        With .TableStyleElements(xlWholeTable).Borders(xlInsideHorizontal)
            .ThemeColor = xlThemeColorAccent1
            .TintAndShade = 0.399975585192419
            .Weight = xlThin
            .LineStyle = xlNone
        End With
        
        ' Configurar fila de encabezado
        With .TableStyleElements(xlHeaderRow).Font
            .FontStyle = "Negrita"
            .TintAndShade = 0
            .ThemeColor = xlThemeColorDark1
        End With
        
        With .TableStyleElements(xlHeaderRow).Interior
            .Pattern = xlSolid
            .PatternThemeColor = xlThemeColorAccent1
            .Color = colorPrimario
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        
        ' Configurar fila total
        With .TableStyleElements(xlTotalRow).Font
            .FontStyle = "Negrita"
            .TintAndShade = 0
            .ThemeColor = xlThemeColorLight1
        End With
        
        With .TableStyleElements(xlTotalRow).Borders(xlEdgeTop)
            .ThemeColor = xlThemeColorAccent1
            .TintAndShade = 0
            .Weight = xlThick
            .LineStyle = xlContinuous
        End With
        
        ' Configurar primera columna
        With .TableStyleElements(xlFirstColumn).Font
            .FontStyle = "Negrita"
            .TintAndShade = 0
            .ThemeColor = xlThemeColorLight1
        End With
        
        ' Configurar última columna
        With .TableStyleElements(xlLastColumn).Font
            .FontStyle = "Negrita"
            .TintAndShade = 0
            .ThemeColor = xlThemeColorLight1
        End With
        
        ' Configurar rayas de filas y columnas
        With .TableStyleElements(xlRowStripe1).Interior
            .Pattern = xlSolid
            .PatternThemeColor = xlThemeColorAccent1
            .ThemeColor = xlThemeColorAccent1
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0.799981688894314
        End With
        
        With .TableStyleElements(xlColumnStripe1).Interior
            .Pattern = xlSolid
            .PatternThemeColor = xlThemeColorAccent1
            .ThemeColor = xlThemeColorAccent1
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0.799981688894314
        End With
    End With
End Sub

