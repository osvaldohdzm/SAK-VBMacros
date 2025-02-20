Attribute VB_Name = "M�dulo1"
Sub GenerarDiasLaborales()
    Dim fechaInicial As String
    Dim fecha As Date
    Dim celda As Range
    Dim fechaFinal As Date
    
    ' Solicitar la fecha inicial
    fechaInicial = InputBox("Introduce la fecha inicial (DD/MM/AAAA):")
    
    ' Verificar si el formato es correcto
    On Error Resume Next
    fecha = DateValue(fechaInicial)
    On Error GoTo 0
    
    ' Si la fecha no es v�lida, mostrar mensaje de error
    If fecha = 0 Then
        MsgBox "Fecha no v�lida, por favor ingresa una fecha en formato DD/MM/AAAA."
        Exit Sub
    End If
    
    ' Recorremos las celdas seleccionadas
    For Each celda In Selection
        ' Si la celda est� vac�a, asignamos la fecha y pasamos al siguiente d�a laboral
        If celda.Value = "" Then
            celda.Value = fecha
            ' Avanzamos al siguiente d�a laboral
            Do
                fecha = fecha + 1
            Loop While Weekday(fecha, vbMonday) > 5 ' Si es fin de semana, buscamos el siguiente lunes
        End If
    Next celda
End Sub

