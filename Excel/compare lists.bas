Attribute VB_Name = "MÃ³dulo1"
Sub CompareLists()
    Dim ws As Worksheet
    Dim rangeA As Range, rangeB As Range
    Dim item As Variant
    Dim onlyInA As Collection, onlyInB As Collection, inBoth As Collection
    Dim outputRow As Long

    ' Selección de la lista A
    On Error Resume Next
    Set rangeA = Application.InputBox("Selecciona el rango para la Lista A", Type:=8)
    If rangeA Is Nothing Then Exit Sub
    
    ' Selección de la lista B
    Set rangeB = Application.InputBox("Selecciona el rango para la Lista B", Type:=8)
    If rangeB Is Nothing Then Exit Sub
    On Error GoTo 0

    ' Inicializar colecciones para almacenar los resultados
    Set onlyInA = New Collection
    Set onlyInB = New Collection
    Set inBoth = New Collection

    ' Crear una hoja para los resultados
    Set ws = Worksheets.Add
    ws.Name = "Resultados Comparación"
    
    ' Recorrer la lista A y clasificar los valores
    For Each item In rangeA.Value
        On Error Resume Next
        If Application.WorksheetFunction.CountIf(rangeB, item) = 0 Then
            onlyInA.Add item, CStr(item)
        Else
            inBoth.Add item, CStr(item)
        End If
        On Error GoTo 0
    Next item
    
    ' Recorrer la lista B para encontrar elementos únicos
    For Each item In rangeB.Value
        On Error Resume Next
        If Application.WorksheetFunction.CountIf(rangeA, item) = 0 Then
            onlyInB.Add item, CStr(item)
        End If
        On Error GoTo 0
    Next item

    ' Escribir resultados en la hoja de cálculo
    ws.Cells(1, 1).Value = "Only in List A"
    ws.Cells(1, 2).Value = "Only in List B"
    ws.Cells(1, 3).Value = "In Both Lists"
    outputRow = 2
    
    ' Escribir la lista de "Only in List A"
    For Each item In onlyInA
        ws.Cells(outputRow, 1).Value = item
        outputRow = outputRow + 1
    Next item

    ' Escribir la lista de "Only in List B"
    outputRow = 2
    For Each item In onlyInB
        ws.Cells(outputRow, 2).Value = item
        outputRow = outputRow + 1
    Next item

    ' Escribir la lista de "In Both Lists"
    outputRow = 2
    For Each item In inBoth
        ws.Cells(outputRow, 3).Value = item
        outputRow = outputRow + 1
    Next item

    ' Agregar botones para copiar al portapapeles
    AddCopyButton ws, "CopyListA", "Only in List A", 1
    AddCopyButton ws, "CopyListB", "Only in List B", 2
    AddCopyButton ws, "CopyBothLists", "In Both Lists", 3

End Sub

Sub AddCopyButton(ws As Worksheet, buttonName As String, listName As String, col As Integer)
    Dim btn As Button
    Set btn = ws.Buttons.Add(Cells(1, col).Left, Cells(1, col).Top, 100, 20)
    btn.Caption = "Copy " & listName
    btn.OnAction = "'" & ThisWorkbook.Name & "'!" & buttonName
    btn.Name = buttonName
End Sub

Sub CopyListA()
    CopyRangeToClipboard ActiveSheet.Range("A2", ActiveSheet.Cells(Rows.Count, 1).End(xlUp))
End Sub

Sub CopyListB()
    CopyRangeToClipboard ActiveSheet.Range("B2", ActiveSheet.Cells(Rows.Count, 2).End(xlUp))
End Sub

Sub CopyBothLists()
    CopyRangeToClipboard ActiveSheet.Range("C2", ActiveSheet.Cells(Rows.Count, 3).End(xlUp))
End Sub

Sub CopyRangeToClipboard(rng As Range)
    Dim DataObj As Object
    Set DataObj = CreateObject("MSForms.DataObject")
    DataObj.SetText Join(Application.Transpose(rng.Value), vbCrLf)
    DataObj.PutInClipboard
    MsgBox "Lista copiada al portapapeles.", vbInformation
End Sub

