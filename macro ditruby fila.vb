Sub DistribuirFilas()

    Dim fila As Long
    Dim valores() As String
    Dim celda As Range
    Dim rango As Range
    Dim nFilas As Long
    Dim columna As String
    
    'Solicitar al usuario la columna para distribuir los valores
    columna = InputBox("Ingrese la letra de la columna para distribuir los valores (por ejemplo, D):")
    
    'Comprobar que el usuario ingresó una letra de columna válida
    If Not IsNumeric(Application.Match(columna, Range("A1:XFD1"), 0)) Then
        MsgBox "La letra de columna ingresada no es válida.", vbExclamation, "Error"
        Exit Sub
    End If
    
    'Definir el rango de celdas para la columna seleccionada
    Set rango = Range(columna & "2:" & columna & Range(columna & Rows.Count).End(xlUp).Row)
    
    For Each celda In rango
    
        'Comprobar si la celda contiene una cadena de valores concatenados por comas
        If InStr(1, celda.Value, ",") > 0 Then
            
            'Separar los valores por comas y guardarlos en un array
            valores = Split(celda.Value, ",")
            nFilas = UBound(valores)
            
            'Insertar filas adicionales
            If nFilas > 0 Then
                celda.Offset(1).Resize(nFilas).EntireRow.Insert Shift:=xlDown
            End If
            
            'Copiar los valores de las otras columnas en filas adicionales
            For i = 0 To nFilas
                
                fila = celda.Row + i
                
                Range("A" & fila & ":Q" & fila).Value = Range("A" & celda.Row & ":Q" & celda.Row).Value
                Range(columna & fila).Value = valores(i)
                
            Next i
            
        End If
    
    Next celda

End Sub
