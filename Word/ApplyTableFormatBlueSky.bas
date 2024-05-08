Attribute VB_Name = "Módulo1"
Sub FormatearTabla()
    Dim tbl As Table
    Dim row As row
    Dim isFirstRow As Boolean
    
    ' Comprobar si hay al menos una tabla seleccionada
    If Selection.Tables.Count = 0 Then
        MsgBox "No hay ninguna tabla seleccionada.", vbExclamation
        Exit Sub
    End If
    
    ' Iterar sobre cada tabla seleccionada
    For Each tbl In Selection.Tables
        ' Formatear la tabla con borde de 1/2 punto y color RGB(0, 112, 192)
        With tbl
            .Borders.Enable = True
            .Borders.InsideLineStyle = wdLineStyleSingle
            .Borders.OutsideLineStyle = wdLineStyleSingle
            .Borders.OutsideLineWidth = wdLineWidth050pt
            .Borders.InsideColor = RGB(0, 112, 192)
            .Borders.OutsideColor = RGB(0, 112, 192)
        End With
        
        ' Formatear la primera fila (cabecera)
        Set row = tbl.Rows(1)
        row.Range.Shading.BackgroundPatternColor = RGB(0, 112, 192)
        row.Range.Font.Color = RGB(255, 255, 255)
        row.Range.ParagraphFormat.SpaceBefore = 0
        row.Range.ParagraphFormat.SpaceAfter = 0
        row.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter 'Centrar texto
        
        ' Iterar sobre las filas restantes
        isFirstRow = True
        For Each row In tbl.Rows
            If Not isFirstRow Then
                ' Aplicar formato a filas no cabecera
                Dim cell As cell
                For Each cell In row.Cells
                    ' Verificar si la celda tiene un color de fondo predeterminado (blanco)
                    If cell.Range.Shading.BackgroundPatternColorIndex = wdColorAutomatic Then
                        If row.Index Mod 2 = 0 Then
                            ' Filas pares
                            row.Range.Shading.BackgroundPatternColor = RGB(255, 255, 255)
                            row.Range.Font.Color = RGB(0, 0, 0) ' Color de letra negro
                        Else
                            ' Filas impares
                            row.Range.Shading.BackgroundPatternColor = RGB(217, 226, 243) ' Color de fondo blanco
                            row.Range.Font.Color = RGB(0, 0, 0) ' Color de letra negro
                        End If
                    End If
                Next cell
                row.Range.ParagraphFormat.SpaceBefore = 0
                row.Range.ParagraphFormat.SpaceAfter = 0
                row.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter 'Centrar texto
                
                ' Centrar contenido verticalmente
                row.Cells.VerticalAlignment = wdCellAlignVerticalCenter
            End If
            isFirstRow = False
        Next row
    Next tbl
    
    MsgBox "La tabla seleccionada ha sido formateada correctamente.", vbInformation
End Sub

