Attribute VB_Name = "Módulo1"
Sub ProcesarCeldas()
    Dim celda As Range
    Dim rangoSeleccion As Range
    Dim textoOriginal As String
    Dim textoLimpiado As String
    Dim datos() As String
    Dim i As Long
    Dim fila As Long
    Dim rng As Range
    Dim dict As Object
    Dim clave As Variant

    ' Selecciona el rango actual
    Set rangoSeleccion = Selection

    Application.ScreenUpdating = False

    ' Paso 1: Eliminar espacios y tabulaciones en las celdas seleccionadas
    For Each celda In rangoSeleccion
        If Not IsEmpty(celda.Value) Then
            ' Elimina espacios y tabulaciones
            celda.Value = Application.WorksheetFunction.Trim(celda.Value)
        End If
    Next celda

    ' Paso 2: Eliminar el texto a partir de los dos puntos, incluyendo los dos puntos
    For Each celda In rangoSeleccion
        If Not IsEmpty(celda.Value) Then
            textoOriginal = celda.Value
            ' Quita todo texto a partir de los dos puntos, incluyendo los dos puntos
            If InStr(textoOriginal, ":") > 0 Then
                textoLimpiado = Left(textoOriginal, InStr(textoOriginal, ":") - 1)
                celda.Value = textoLimpiado
            End If
        End If
    Next celda

    ' Paso 3: Dividir las celdas con saltos de línea y crear nuevas filas
    fila = 1
    Do While fila <= ActiveSheet.UsedRange.Rows.Count
        If Not IsEmpty(Cells(fila, 1).Value) And Not IsEmpty(Cells(fila, 2).Value) Then
            datos = Split(Cells(fila, 1).Value, Chr(10))
            If UBound(datos) > 0 Then
                ' Si hay múltiples valores en la celda
                For i = UBound(datos) To 0 Step -1
                    ' Inserta una nueva fila
                    Rows(fila + 1).Insert Shift:=xlDown
                    ' Copia el valor de la celda a la nueva fila
                    Cells(fila + 1, 1).Value = datos(i)
                    Cells(fila + 1, 2).Value = Cells(fila, 2).Value
                Next i
                ' Borra el valor original
                Cells(fila, 1).Value = datos(0)
            End If
        End If
        fila = fila + 1
    Loop

    Set rng = rangoSeleccion
    rng.RemoveDuplicates Columns:=1, Header:= _
        xlYes

    Application.ScreenUpdating = True
End Sub

