Attribute VB_Name = "Módulo2"
Sub CombinarRegistrosEnHojaActual()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim selectedColumn As Range
    Dim dictApp As Object
    Dim dictDB As Object
    Dim cell As Range
    Dim key As Variant
    
    ' Definir la hoja de trabajo actual
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Definir la columna seleccionada por el usuario
    On Error Resume Next
    Set selectedColumn = Application.InputBox("Selecciona la columna a procesar:", Type:=8)
    On Error GoTo 0
    
    ' Salir si no se selecciona ninguna columna
    If selectedColumn Is Nothing Then
        Exit Sub
    End If
    
    ' Inicializar diccionarios para almacenar valores únicos
    Set dictApp = CreateObject("Scripting.Dictionary")
    Set dictDB = CreateObject("Scripting.Dictionary")
    
    ' Obtener la última fila en la columna seleccionada
    lastRow = ws.Cells(ws.Rows.Count, selectedColumn.Column).End(xlUp).Row
    
    ' Recorrer los registros y combinar los valores
    For Each cell In ws.Range(selectedColumn.Cells(2, 1), selectedColumn.Cells(lastRow, 1))
        ' Combinar valores de AssociatedApplication
        If Not dictApp.Exists(cell.Value) Then
            dictApp.Add cell.Value, cell.Offset(0, 4).Value
        Else
            dictApp(cell.Value) = dictApp(cell.Value) & "," & cell.Offset(0, 4).Value
        End If
        
        ' Combinar valores de AssociatedDatabase
        If Not dictDB.Exists(cell.Value) Then
            dictDB.Add cell.Value, cell.Offset(0, 5).Value
        Else
            dictDB(cell.Value) = dictDB(cell.Value) & "," & cell.Offset(0, 5).Value
        End If
    Next cell
    
    ' Actualizar los valores en las celdas
    For Each key In dictApp.keys
        ws.Range(selectedColumn.Cells(2, 1), selectedColumn.Cells(lastRow, 1)).AutoFilter Field:=1, Criteria1:=key
        ws.Cells(ws.Rows.Count, 1).End(xlUp).Offset(1, 4).Value = dictApp(key)
        ws.Cells(ws.Rows.Count, 1).End(xlUp).Offset(1, 5).Value = dictDB(key)
        ws.ShowAllData
    Next key
    
    ' Limpiar diccionarios
    Set dictApp = Nothing
    Set dictDB = Nothing
End Sub

