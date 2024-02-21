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
        ' Formatear la primera fila (cabecera)
        Set row = tbl.Rows(1)
        row.Range.Shading.BackgroundPatternColor = RGB(0, 32, 96) ' Color de fondo #002060
        row.Range.Font.Color = RGB(255, 255, 255) ' Color de letra blanco
        row.Range.ParagraphFormat.SpaceBefore = 0
        row.Range.ParagraphFormat.SpaceAfter = 0
        row.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter 'Centrar texto
        
        ' Iterar sobre las filas restantes
        isFirstRow = True
        For Each row In tbl.Rows
            If Not isFirstRow Then
                ' Aplicar formato a filas no cabecera
                If row.Index Mod 2 = 0 Then
                    ' Filas pares
                    row.Range.Shading.BackgroundPatternColor = RGB(217, 225, 242) ' Azul D9E1F2
                    row.Range.Font.Color = RGB(0, 0, 0) ' Color de letra negro
                Else
                    ' Filas impares
                    row.Range.Shading.BackgroundPatternColor = RGB(255, 255, 255) ' Color de fondo blanco
                    row.Range.Font.Color = RGB(0, 0, 0) ' Color de letra negro
                End If
                row.Range.ParagraphFormat.SpaceBefore = 0
                row.Range.ParagraphFormat.SpaceAfter = 0
                row.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter 'Centrar texto
            End If
            isFirstRow = False
        Next row
    Next tbl
    
    MsgBox "La tabla seleccionada ha sido formateada correctamente.", vbInformation
End Sub
