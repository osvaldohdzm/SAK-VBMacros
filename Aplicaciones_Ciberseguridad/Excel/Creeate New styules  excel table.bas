Attribute VB_Name = "Módulo1"
Sub CrearEstiloTablaPersonalizado()
    Dim libro As Workbook
    Dim ts As TableStyle
    Dim estiloNombre As String
    Dim tabla As ListObject

    ' Define el libro actual
    Set libro = ThisWorkbook

    ' Nombre del estilo de tabla que deseas crear
    estiloNombre = "EstiloPersonalizado"

    ' Elimina el estilo existente con el mismo nombre si ya existe
    On Error Resume Next
    libro.TableStyles(estiloNombre).Delete
    On Error GoTo 0

    ' Crea un nuevo estilo de tabla personalizado
    Set ts = libro.TableStyles.Add(estiloNombre)

    ' Configura las opciones del estilo de tabla
    With ts
        .ShowAsAvailablePivotTableStyle = False
        .ShowAsAvailableTableStyle = True
        .ShowAsAvailableSlicerStyle = False
        .ShowAsAvailableTimelineStyle = False
    End With

    ' Configura el estilo del encabezado de la tabla
    DoFontHeader ts.TableStyleElements(xlHeaderRow)
    DoInteriorHeader ts.TableStyleElements(xlHeaderRow)

    ' Configura el estilo de las filas alternas (efecto zebra)
    DoFontRow ts.TableStyleElements(xlRowStripe1), True
    DoFontRow ts.TableStyleElements(xlRowStripe2), False

    ' Configura el estilo de la fila totalizadora si está presente
    DoFontTotalRow ts.TableStyleElements(xlTotalRow)

    ' Aplica el estilo a una tabla existente llamada "MiTabla"
    On Error Resume Next
    Set tabla = ActiveSheet.ListObjects("MiTabla")
    On Error GoTo 0

    If Not tabla Is Nothing Then
        tabla.TableStyle = estiloNombre
    End If
End Sub

Sub DoFontHeader(tse As TableStyleElement)
    With tse.Font
        .Color = RGB(255, 255, 255) ' Blanco
        .FontStyle = "Bold"
    End With
End Sub

Sub DoInteriorHeader(tse As TableStyleElement)
    With tse.Interior
        .Color = RGB(0, 0, 128) ' Azul marino
    End With
End Sub

Sub DoFontRow(tse As TableStyleElement, isStripe1 As Boolean)
    With tse.Font
        If isStripe1 Then
            .Color = RGB(255, 255, 255) ' Blanco
        Else
            .Color = RGB(0, 0, 0) ' Negro
        End If
    End With
End Sub

Sub DoFontTotalRow(tse As TableStyleElement)
    With tse.Font
        .Color = RGB(255, 255, 255) ' Blanco
        .FontStyle = "Bold"
    End With
End Sub

