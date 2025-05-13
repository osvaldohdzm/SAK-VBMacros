Attribute VB_Name = "Module3"
Sub CrearTablaINAI()
    Dim wb As Workbook
    Dim tblStyle As TableStyle
    
    ' Referencia al libro activo
    Set wb = ActiveWorkbook
    
    ' Crear el estilo de tabla "Tabla INAI"
    Set tblStyle = wb.TableStyles.Add("Tabla INAI")
    
    ' Establecer propiedades del estilo de tabla
    With tblStyle
        .TableStyleElements(TableConstants.xlHeaderRow).Interior.Color = RGB(31, 78, 121) ' Color azul (#1F4E79) para las cabeceras
        .TableStyleElements(TableConstants.xlSecondRowStripe).Interior.Color = RGB(31, 78, 121) ' Color azul (#1F4E79) para las filas alternas
        
        ' Si se desea modificar otras partes de la tabla, se pueden establecer aquí las propiedades adicionales
        ' Por ejemplo:
        ' .TableStyleElements(TableConstants.xlWholeTable).Borders.LineStyle = xlContinuous
        ' .TableStyleElements(TableConstants.xlFirstColumn).Font.Bold = True
        ' .TableStyleElements(TableConstants.xlTotalRow).Font.Color = RGB(255, 0, 0) ' Ejemplo para la fila total en rojo
    End With
    
    ' Mostrar el estilo de tabla como disponible para tablas normales
    With tblStyle
        .ShowAsAvailablePivotTableStyle = False
        .ShowAsAvailableTableStyle = True
        .ShowAsAvailableSlicerStyle = False
        .ShowAsAvailableTimelineStyle = False
    End With
    
    MsgBox "Se ha creado el estilo de tabla 'Tabla INAI'.", vbInformation
End Sub

