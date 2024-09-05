Attribute VB_Name = "NewMacros"
Sub FormatoNegritaViñetas()
    Dim p As Paragraph
    Dim strTexto As String
    Dim arrLineas As Variant
    Dim i As Integer
    Dim posDosPuntos As Integer
    Dim rng As Range
    
    ' Recorrer cada párrafo en la selección
    For Each p In Selection.Paragraphs
        If p.Range.ListFormat.ListType = wdListBullet Then ' Verificar que sea una viñeta
        
            ' Asignar el texto seleccionado a una cadena
            strTexto = p.Range.Text
            
            ' Dividir el texto en líneas
            arrLineas = Split(strTexto, vbCrLf)
            
            ' Recorrer cada línea
            For i = LBound(arrLineas) To UBound(arrLineas)
                ' Obtener la posición de los dos puntos
                posDosPuntos = InStr(arrLineas(i), ":")
                
                ' Seleccionar el texto desde el primer carácter hasta posDosPuntos y aplicar negrita
                Set rng = p.Range
                rng.MoveStart unit:=wdCharacter, Count:=0
                rng.MoveEnd unit:=wdCharacter, Count:=posDosPuntos - 1
                rng.Font.Bold = True
                
                ' Seleccionar el texto desde posDosPuntos hasta el final y quitar negrita
                rng.MoveStart unit:=wdCharacter, Count:=posDosPuntos - 1
                rng.MoveEnd unit:=wdCharacter, Count:=Len(arrLineas(i)) - posDosPuntos + 1
                rng.Font.Bold = False
            Next i
            
        End If
    Next p
End Sub
