Attribute VB_Name = "Módulo2"
Sub EliminarFilasDuplicadasEnSeleccion()
    Dim rng As Range
    Dim dict As Object
    Dim cellValue As String
    Dim i As Long
    
    On Error Resume Next
    Set rng = Selection.SpecialCells(xlCellTypeConstants)
    On Error GoTo 0
    
    If rng Is Nothing Then
        MsgBox "No se ha seleccionado un rango válido.", vbExclamation
        Exit Sub
    End If
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    For i = rng.Rows.Count To 1 Step -1
        cellValue = Trim(rng.Cells(i).Value)
        If cellValue <> "" And Not dict.Exists(cellValue) Then
            dict.Add cellValue, i
        Else
            rng.Rows(i).EntireRow.Delete
        End If
    Next i
    
    Set dict = Nothing
End Sub

