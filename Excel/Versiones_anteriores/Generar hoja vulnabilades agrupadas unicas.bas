Attribute VB_Name = "Módulo2"
Sub GeneraHojaVulnerabilidadesUnicasYAgrupadas()
    ' Declaración de variables
    Dim ws As Worksheet
    Dim rngHeaders As Range
    Dim dictAgrupadas As Object
    Dim dictUnicas As Object
    Dim cell As Range
    Dim key As Variant
    Dim i As Integer
    Dim firstKey As Variant ' Variable para almacenar la primera clave del diccionario de vulnerabilidades agrupadas

    ' Referencia a la hoja activa
    Set ws = ThisWorkbook.ActiveSheet

    ' Selección del rango de encabezados
    On Error Resume Next
    Set rngHeaders = Application.InputBox("Seleccione el rango con los encabezados", Type:=8)
    On Error GoTo 0

    ' Verificación de que se haya seleccionado un rango válido
    If rngHeaders Is Nothing Then
        MsgBox "No se ha seleccionado un rango válido.", vbExclamation
        Exit Sub
    End If

    ' Creación de una nueva hoja para las vulnerabilidades unicas
    Dim newWsUnicas As Worksheet
    Set newWsUnicas = ThisWorkbook.Sheets.Add(After:=ws)
    newWsUnicas.Name = "Vulnerabilidades unicas"

    ' Copiar el formato y los valores de las celdas de encabezado seleccionadas a la nueva hoja de vulnerabilidades unicas
    Dim header As Range
    For Each header In rngHeaders.Cells
        header.Copy newWsUnicas.Cells(1, header.Column)
    Next header

    ' Definir las columnas necesarias en la nueva hoja de vulnerabilidades unicas
    Dim columnsDict As Object
    Set columnsDict = CreateObject("Scripting.Dictionary")

    ' Identificar las columnas por sus nombres en la hoja de vulnerabilidades unicas
    For Each cell In newWsUnicas.Rows(1).SpecialCells(xlCellTypeConstants)
        columnsDict(cell.Value) = cell.Column
    Next cell

    ' Verificar que existan claves en el diccionario antes de acceder a ellas en la hoja de vulnerabilidades unicas
    If columnsDict.Count > 0 Then
        ' Iterar sobre las claves para obtener la primera clave en la hoja de vulnerabilidades unicas
        For Each key In columnsDict.Keys
            firstKey = key
            Exit For ' Salir del bucle una vez que se haya obtenido la primera clave en la hoja de vulnerabilidades unicas
        Next key
    End If

    ' Crear un diccionario para almacenar los datos agrupados en la hoja de vulnerabilidades unicas
    Set dictUnicas = CreateObject("Scripting.Dictionary")

    ' Iterar sobre los datos y agruparlos en la hoja de vulnerabilidades unicas
    For i = 2 To ws.Cells(ws.Rows.Count, columnsDict(firstKey)).End(xlUp).row
        Dim keyStr As String
        keyStr = ""
        For Each cell In rngHeaders
            keyStr = keyStr & "|" & ws.Cells(i, columnsDict(cell.Value)).Value
        Next cell

        ' Construir la clave para el diccionario en la hoja de vulnerabilidades unicas
        key = Mid(keyStr, 2)

        ' Verificar si la clave ya existe en el diccionario en la hoja de vulnerabilidades unicas
        If Not dictUnicas.Exists(key) Then
            ' Si no existe, crear una nueva entrada en el diccionario en la hoja de vulnerabilidades unicas
            Set dictUnicas(key) = CreateObject("Scripting.Dictionary")
            For Each cell In rngHeaders
                dictUnicas(key)(cell.Value) = ws.Cells(i, columnsDict(cell.Value)).Value
            Next cell
        End If
    Next i

    ' Escribir los resultados en la nueva hoja de vulnerabilidades unicas
    Dim outputRowUnicas As Integer
    outputRowUnicas = 2

    ' Iterar sobre el diccionario y escribir los datos en la nueva hoja de vulnerabilidades unicas
    For Each key In dictUnicas.Keys
        For Each cell In rngHeaders
            newWsUnicas.Cells(outputRowUnicas, columnsDict(cell.Value)).Value = dictUnicas(key)(cell.Value)
        Next cell
        outputRowUnicas = outputRowUnicas + 1
    Next key
    
    ' Ajustar la altura máxima de las filas en la hoja de vulnerabilidades unicas
    With newWsUnicas
        For Each cell In .UsedRange
            If Len(cell.Value) > 0 Then
                If cell.RowHeight > 15 Then
                    cell.RowHeight = 15
                End If
            End If
        Next cell
    End With
    
    ' Eliminar las columnas anteriores a la selección de encabezados en la nueva hoja de vulnerabilidades unicas
    If rngHeaders.Column > 1 Then
        newWsUnicas.Columns("A:" & Split(newWsUnicas.Cells(, rngHeaders.Column - 1).Address, "$")(1)).Delete
    End If
    
    ' Convertir el rango de datos en la hoja de vulnerabilidades unicas en una tabla
    Dim tblUnicas As ListObject
    Set tblUnicas = newWsUnicas.ListObjects.Add(xlSrcRange, newWsUnicas.UsedRange, , xlYes)
    tblUnicas.TableStyle = "TableStyleMedium9"


    ' Definir los nombres de las columnas para usar en el método RemoveDuplicates
Dim colSeveridad As String
Dim colNombreVulnerabilidad As String

' Asignar los nombres de las columnas
colSeveridad = "Severidad"
colNombreVulnerabilidad = "NombreVulnerabilidad"

' Declarar una variable para almacenar el índice de la columna
Dim idxSeveridad As Long
Dim idxNombreVulnerabilidad As Long
Dim col As ListColumn

' Iterar sobre las columnas de la tabla
For Each col In tblUnicas.ListColumns
    ' Verificar si el nombre de la columna coincide con colSeveridad
    If col.Name = colSeveridad Then
        ' Asignar el índice de la columna si se encuentra
        idxSeveridad = col.Index
        Exit For ' Salir del bucle una vez que se ha encontrado la columna
    End If
Next col

' Verificar si se encontró la columna de severidad
If idxSeveridad = 0 Then
    MsgBox "La columna de severidad no existe en la tabla.", vbExclamation
    Exit Sub
End If

' Hacer lo mismo para la columna de nombre de vulnerabilidad
For Each col In tblUnicas.ListColumns
    If col.Name = colNombreVulnerabilidad Then
        idxNombreVulnerabilidad = col.Index
        Exit For
    End If
Next col

If idxNombreVulnerabilidad = 0 Then
    MsgBox "La columna de nombre de vulnerabilidad no existe en la tabla.", vbExclamation
    Exit Sub
End If


' Eliminar duplicados
tblUnicas.Range.RemoveDuplicates Columns:=Array(idxSeveridad, idxNombreVulnerabilidad), header:=xlYes

    ' Ahora, vamos a generar la hoja "Vulnerabilidades agrupadas" y luego la hoja "Vulnerabilidades agrupadas unicas"

    ' Crear una nueva hoja para las vulnerabilidades agrupadas
    Dim newWsAgrupadas As Worksheet
    Set newWsAgrupadas = ThisWorkbook.Sheets.Add(After:=newWsUnicas)
    newWsAgrupadas.Name = "Vulnerabilidades agrupadas"

    ' Copiar el formato y los valores de las celdas de encabezado seleccionadas a la nueva hoja de vulnerabilidades agrupadas
    For Each header In rngHeaders.Cells
        header.Copy newWsAgrupadas.Cells(1, header.Column)
    Next header

    ' Definir las columnas necesarias en la nueva hoja de vulnerabilidades agrupadas
    Dim severityColumn As Integer
    Dim vulnerabilityColumn As Integer
    Dim pathColumn As Integer
    Dim secTestOutputColumn As Integer

    ' Identificar las columnas por sus nombres en la hoja de vulnerabilidades agrupadas
    For Each cell In newWsAgrupadas.Rows(1).SpecialCells(xlCellTypeConstants)
        Select Case cell.Value
            Case "Severidad"
                severityColumn = cell.Column
            Case "NombreVulnerabilidad"
                vulnerabilityColumn = cell.Column
            Case "Ruta"
                pathColumn = cell.Column
            Case "SecTestOutput"
                secTestOutputColumn = cell.Column
        End Select
    Next cell

    ' Si no se encuentran todas las columnas necesarias, salir del subproceso
    If severityColumn = 0 Or vulnerabilityColumn = 0 Or pathColumn = 0 Or secTestOutputColumn = 0 Then
        MsgBox "No se encontraron todas las columnas necesarias en los encabezados.", vbExclamation
        Exit Sub
    End If

    ' Crear un diccionario para almacenar los datos agrupados en la hoja de vulnerabilidades agrupadas
    Set dictAgrupadas = CreateObject("Scripting.Dictionary")

    ' Iterar sobre los datos y agruparlos en la hoja de vulnerabilidades agrupadas
    For i = 2 To ws.Cells(ws.Rows.Count, severityColumn).End(xlUp).row
        Dim severity As String
        Dim vulnerability As String
        Dim ruta As String
        Dim secTestOutput As String
        Dim otherData As String

        severity = ws.Cells(i, severityColumn).Value
        vulnerability = ws.Cells(i, vulnerabilityColumn).Value
        ruta = ws.Cells(i, pathColumn).Value
        secTestOutput = ws.Cells(i, secTestOutputColumn).Value

        ' Construir la clave para el diccionario en la hoja de vulnerabilidades agrupadas
        key = severity & "|" & vulnerability

        ' Verificar si la clave ya existe en el diccionario en la hoja de vulnerabilidades agrupadas
        If dictAgrupadas.Exists(key) Then
            ' Si existe, agregar las rutas y los secTestOutputs a las entradas existentes en la hoja de vulnerabilidades agrupadas
            dictAgrupadas(key)("Ruta") = dictAgrupadas(key)("Ruta") & vbCrLf & ruta
            If Len(secTestOutput) > 0 Then
                dictAgrupadas(key)("SecTestOutput") = dictAgrupadas(key)("SecTestOutput") & vbCrLf & vbCrLf & ruta & " ------>" & vbCrLf & secTestOutput
            End If
        Else
            ' Si no existe, crear una nueva entrada en el diccionario en la hoja de vulnerabilidades agrupadas
            Set dictAgrupadas(key) = CreateObject("Scripting.Dictionary")
            dictAgrupadas(key)("Severidad") = severity
            dictAgrupadas(key)("NombreVulnerabilidad") = vulnerability
            dictAgrupadas(key)("Ruta") = ruta
            dictAgrupadas(key)("SecTestOutput") = ruta & " ------>" & vbCrLf & secTestOutput
        End If
    Next i

    ' Escribir los resultados en la nueva hoja de vulnerabilidades agrupadas
    Dim outputRowAgrupadas As Integer
    outputRowAgrupadas = 2

    ' Iterar sobre el diccionario y escribir los datos en la nueva hoja de vulnerabilidades agrupadas
    For Each key In dictAgrupadas.Keys
        newWsAgrupadas.Cells(outputRowAgrupadas, severityColumn).Value = dictAgrupadas(key)("Severidad")
        newWsAgrupadas.Cells(outputRowAgrupadas, vulnerabilityColumn).Value = dictAgrupadas(key)("NombreVulnerabilidad")
        newWsAgrupadas.Cells(outputRowAgrupadas, pathColumn).Value = dictAgrupadas(key)("Ruta")
        newWsAgrupadas.Cells(outputRowAgrupadas, secTestOutputColumn).Value = dictAgrupadas(key)("SecTestOutput")
        outputRowAgrupadas = outputRowAgrupadas + 1
    Next key
    
    ' Ajustar la altura máxima de las filas en la hoja de vulnerabilidades agrupadas
    With newWsAgrupadas
        For Each cell In .UsedRange
            If Len(cell.Value) > 0 Then
                If cell.RowHeight > 15 Then
                    cell.RowHeight = 15
                End If
            End If
        Next cell
    End With
    
    ' Eliminar las columnas anteriores a la selección de encabezados en la nueva hoja de vulnerabilidades agrupadas
    If rngHeaders.Column > 1 Then
        newWsAgrupadas.Columns("A:" & Split(newWsAgrupadas.Cells(, rngHeaders.Column - 1).Address, "$")(1)).Delete
    End If

     ' Convertir el rango de datos en la hoja de vulnerabilidades unicas en una tabla
    Dim tblAgrupadas As ListObject
    Set tblAgrupadas = newWsAgrupadas.ListObjects.Add(xlSrcRange, newWsAgrupadas.UsedRange, , xlYes)
    tblAgrupadas.TableStyle = "TableStyleMedium9"

    ' Generar la hoja "Vulnerabilidades agrupadas unicas"
    ' Utilizando los datos de la hoja de vulnerabilidades agrupadas y los valores de la hoja de vulnerabilidades unicas

  ' Duplicar la hoja "Vulnerabilidades unicas"
    newWsUnicas.Copy After:=newWsAgrupadas
    ActiveSheet.Name = "Vulns agrupadas_unicas"
    
    ' Obtener una referencia a la hoja de "Vulns agrupadas_unicas"
    Dim newWsAgrupadasUnicas As Worksheet
    Set newWsAgrupadasUnicas = ActiveSheet

        ' Definir las hojas de trabajo
    Dim wsVulnsAgrupadasUnicas As Worksheet
    Dim wsVulnerabilidadesAgrupadas As Worksheet
    
    ' Asignar las hojas de trabajo
    Set wsVulnsAgrupadasUnicas = ThisWorkbook.Sheets("Vulns agrupadas_unicas")
    Set wsVulnerabilidadesAgrupadas = ThisWorkbook.Sheets("Vulnerabilidades agrupadas")
    
    Dim colSeveridadNum As Integer ' Declaración de la variable colSeveridadNum como tipo Integer
colSeveridadNum = 1 ' Asignación de un valor a la variable colSeveridadNum

 Dim colNombreVulnerabilidadNum As Integer ' Declaración de la variable colSeveridadNum como tipo Integer
colNombreVulnerabilidadNum = 1 ' Asignación de un valor a la variable colSeveridadNum

    ' Buscar los índices de columna para los encabezados relevantes en "Vulns agrupadas_unicas"
    colSeveridadNum = GetColumnIndex(wsVulnsAgrupadasUnicas, "Severidad")
    colNombreVulnerabilidadNum = GetColumnIndex(wsVulnsAgrupadasUnicas, "NombreVulnerabilidad")
    colSecTestOutput = GetColumnIndex(wsVulnsAgrupadasUnicas, "SecTestOutput")
    colRuta = GetColumnIndex(wsVulnsAgrupadasUnicas, "Ruta")
    
    ' Iterar sobre cada fila en "Vulns agrupadas_unicas"
    Dim l As Long
    For l = 2 To wsVulnsAgrupadasUnicas.Cells(wsVulnsAgrupadasUnicas.Rows.Count, colSeveridadNum).End(xlUp).row
        ' Obtener el valor de Severidad y NombreVulnerabilidad para buscar en "Vulnerabilidades agrupadas"
        Dim severidad As String
        Dim nombreVulnerabilidad As String
        severidad = wsVulnsAgrupadasUnicas.Cells(l, colSeveridadNum).Value
        nombreVulnerabilidad = wsVulnsAgrupadasUnicas.Cells(l, colNombreVulnerabilidadNum).Value
        
        ' Buscar la coincidencia en "Vulnerabilidades agrupadas"
        Dim searchResult As Range
        Set searchResult = wsVulnerabilidadesAgrupadas.Range("A:A").Find(What:=severidad, LookIn:=xlValues, LookAt:=xlWhole)
        
        ' Si se encuentra la coincidencia
        If Not searchResult Is Nothing Then
            Do
                ' Verificar si el nombre de la vulnerabilidad coincide
                If searchResult.Offset(0, 1).Value = nombreVulnerabilidad Then
                    ' Copiar los valores de SecTestOutput y Ruta
                    wsVulnsAgrupadasUnicas.Cells(l, colSecTestOutput).Value = searchResult.Offset(0, colSecTestOutput - 1).Value ' SecTestOutput
                    wsVulnsAgrupadasUnicas.Cells(l, colRuta).Value = searchResult.Offset(0, colRuta - 1).Value ' Ruta
                    Exit Do
                End If
                ' Buscar la siguiente coincidencia en "Vulnerabilidades agrupadas"
                Set searchResult = wsVulnerabilidadesAgrupadas.Range("A:A").FindNext(searchResult)
            Loop While Not searchResult Is Nothing And searchResult.row <> wsVulnerabilidadesAgrupadas.Range("A:A").Find(What:=severidad, After:=searchResult, LookIn:=xlValues, LookAt:=xlWhole).row
        End If
    Next l


' Obtener el índice de la columna SecTestOutput en "Vulns agrupadas_unicas"
colSecTestOutput = GetColumnIndex(wsVulnsAgrupadasUnicas, "SecTestOutput")

' Verificar si la columna "SecTestOutput" existe en la hoja "Vulns agrupadas_unicas"
If colSecTestOutput > 0 Then
    ' Aplicar el primer código para limpiar el contenido de las celdas en la columna "SecTestOutput"
    Dim celda As Range
    For Each celda In wsVulnsAgrupadasUnicas.Columns(colSecTestOutput).SpecialCells(xlCellTypeConstants)
        Dim lineas() As String
        Dim indexn As Integer
        ' Dividir el contenido de la celda en líneas
        lineas = Split(celda.Value, vbCrLf)
        ' Recorrer cada línea y eliminar los espacios en blanco a la izquierda
        For indexn = LBound(lineas) To UBound(lineas)
            lineas(indexn) = Trim(Replace(lineas(indexn), Chr(9), ""))
        Next indexn
        ' Actualizar el contenido de la celda con el texto limpio
        celda.Value = Join(lineas, vbCrLf)
    Next celda

    ' Aplicar el segundo código para eliminar líneas vacías y ajustar el formato
    For Each celda In wsVulnsAgrupadasUnicas.Columns(colSecTestOutput).SpecialCells(xlCellTypeConstants)
        ' Reemplazar diferentes saltos de línea con vbLf
        Dim contenido As String
        contenido = Replace(Replace(Replace(celda.Value, vbCrLf, vbLf), vbCr, vbLf), vbLf & vbLf, vbLf)
        ' Si el contenido comienza con vbLf, quitarlo
        If Left(contenido, 1) = vbLf Then
            contenido = Mid(contenido, 2)
        End If
        ' Si el contenido termina con vbLf, quitarlo
        If Right(contenido, 1) = vbLf Then
            contenido = Left(contenido, Len(contenido) - 1)
        End If
        ' Dividir el contenido de la celda en un array de líneas
        lineas = Split(contenido, vbLf)
        ' Iterar sobre cada línea del array
        For indexn = LBound(lineas) To UBound(lineas)
            ' Verificar si la línea está vacía y eliminarla
            If Trim(lineas(indexn)) = "" Then
                lineas(indexn) = vbNullString
            End If
        Next indexn
        ' Unir el array de líneas de nuevo en una cadena y asignarlo a la celda
        celda.Value = Join(lineas, vbLf)
    Next celda
Else
    MsgBox "La columna 'SecTestOutput' no se encontró en la hoja 'Vulns agrupadas_unicas'.", vbExclamation
End If
    
    ' Eliminar las hojas "Vulnerabilidades unicas" y "Vulnerabilidades agrupadas"
On Error Resume Next ' Ignorar errores si las hojas no existen
Application.DisplayAlerts = False ' Desactivar las alertas de eliminación
ThisWorkbook.Sheets("Vulnerabilidades unicas").Delete
ThisWorkbook.Sheets("Vulnerabilidades agrupadas").Delete
Application.DisplayAlerts = True ' Activar nuevamente las alertas
On Error GoTo 0 ' Volver a habilitar el manejo de errores

    MsgBox "Proceso completado. Los registros homologados se han creado en las hojas 'Vulnerabilidades unicas', 'Vulnerabilidades agrupadas' y 'Vulnerabilidades agrupadas unicas'."
End Sub

Function GetColumnIndex(ws As Worksheet, header As String) As Long
    ' Buscar el encabezado y devolver su índice de columna
    Dim headerCell As Range
    Set headerCell = ws.Rows(1).Find(What:=header, LookIn:=xlValues, LookAt:=xlWhole)
    If Not headerCell Is Nothing Then
        GetColumnIndex = headerCell.Column
    Else
        GetColumnIndex = 0 ' Devolver 0 si no se encuentra el encabezado
    End If
End Function
