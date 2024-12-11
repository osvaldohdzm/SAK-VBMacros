Attribute VB_Name = "Módulo2"
Sub DividirIPs()
    Dim ws As Worksheet
    Dim selectedRange As Range
    Dim cell As Range
    Dim ipArray() As String
    Dim newRow As Long
    Dim i As Long, j As Long
    
    ' Definir la hoja de trabajo
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Verificar si hay solo una columna seleccionada
    If Selection.Columns.Count > 1 Then
        MsgBox "Selecciona solo una columna que contenga las direcciones IP separadas por comas.", vbExclamation
        Exit Sub
    End If
    
    ' Iterar sobre cada celda en la selección
    For Each cell In Selection
        ' Verificar si hay múltiples direcciones IP en la celda
        If InStr(cell.Value, ",") > 0 Then
            ' Dividir las direcciones IP en un array
            ipArray = Split(cell.Value, ",")
            
            ' Obtener la fila actual
            newRow = cell.Row
            
            ' Insertar una nueva fila para cada dirección IP adicional
            For j = LBound(ipArray) To UBound(ipArray)
                ' Insertar una nueva fila debajo de la fila actual
                ws.Rows(newRow + 1).Insert Shift:=xlDown
                
                ' Copiar el contenido de la fila actual a la nueva fila
                ws.Rows(newRow).Copy Destination:=ws.Rows(newRow + 1)
                
                ' Pegar la dirección IP en la nueva fila
                ws.Cells(newRow + 1, cell.Column).Value = Trim(ipArray(j))
                
                ' Actualizar el número de fila actual
                newRow = newRow + 1
            Next j
            
            ' Eliminar la dirección IP original de la celda
            cell.Value = Trim(ipArray(0))
        End If
    Next cell
End Sub

