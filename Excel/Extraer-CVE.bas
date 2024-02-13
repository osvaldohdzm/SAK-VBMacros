Attribute VB_Name = "Módulo1"
Sub ExtraerCVEEnFila()
    ' Obtiene las celdas seleccionadas
    Dim celdasSeleccionadas As Range
    Set celdasSeleccionadas = Selection

    ' Inserta una columna adyacente a la derecha de las celdas seleccionadas
    celdasSeleccionadas.EntireColumn.Insert Shift:=xlShiftToLeft

    ' Recorre cada celda en las celdas seleccionadas
    Dim celda As Range
    For Each celda In celdasSeleccionadas
        ' Busca CVE en el contenido de la celda
        Dim inicioCVE As Integer
        inicioCVE = InStr(celda.Value, "CVE-")

        ' Si se encuentra CVE en la celda
        If inicioCVE > 0 Then
            ' Encuentra la posición de la coma después de CVE
            Dim finCVE As Integer
            finCVE = InStr(inicioCVE, celda.Value, ",")

            ' Si no hay coma después de CVE
            If finCVE = 0 Then
                ' Extrae el CVE desde el inicio de CVE hasta el final de la celda
                celda.Offset(0, -1).Value = Mid(celda.Value, inicioCVE)
            Else
                ' Extrae el CVE desde el inicio de CVE hasta la coma antes del siguiente CVE
                celda.Offset(0, -1).Value = Mid(celda.Value, inicioCVE, finCVE - inicioCVE)
            End If
        Else
            ' Si no se encuentra CVE, coloca "No tiene" en la columna adyacente
            celda.Offset(0, -1).Value = ""
        End If
    Next celda
End Sub

