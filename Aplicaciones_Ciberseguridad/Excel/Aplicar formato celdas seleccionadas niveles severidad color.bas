Attribute VB_Name = "Módulo11"
Sub AplicarFormatoPorContenido()
    Dim selectedRange As Range
    Dim cell As Range

    ' Verificar si hay celdas seleccionadas
    On Error Resume Next
    Set selectedRange = Selection.SpecialCells(xlCellTypeConstants)
    On Error GoTo 0

    ' Realizar los reemplazos necesarios
    If Not selectedRange Is Nothing Then
        For Each cell In selectedRange
            Select Case UCase(cell.Value)
                Case "BAJO"
                    cell.Value = "BAJA"
                Case "MEDIO"
                    cell.Value = "MEDIA"
                Case "ALTO"
                    cell.Value = "ALTA"
                Case "CRITICO"
                    cell.Value = "CRÍTICA"
                Case "INFOMATIVO"
                    cell.Value = "INFORMATIVA"
            End Select
        Next cell
    End If

    ' Aplicar formato según el contenido de las celdas seleccionadas
    For Each cell In selectedRange
        Select Case UCase(cell.Value)
            Case "CRÍTICA"
                cell.Font.Color = RGB(255, 255, 255) ' Blanco
                cell.Interior.Color = RGB(112, 48, 160) ' #7030A0
            Case "ALTA"
                cell.Font.Color = RGB(255, 255, 255) ' Blanco
                cell.Interior.Color = RGB(255, 0, 0) ' #FF0000
            Case "MEDIA"
                cell.Font.Color = RGB(0, 0, 0) ' Negro
                cell.Interior.Color = RGB(255, 255, 0) ' #FFFF00
            Case "BAJA"
                cell.Font.Color = RGB(255, 255, 255) ' Blanco
                cell.Interior.Color = RGB(0, 176, 80) ' #00B050
            Case "INFORMATIVA"
                cell.Font.Color = RGB(0, 0, 0) ' Negro
                cell.Interior.Color = RGB(231, 230, 230) ' #E7E6E6
        End Select
    Next cell
End Sub

