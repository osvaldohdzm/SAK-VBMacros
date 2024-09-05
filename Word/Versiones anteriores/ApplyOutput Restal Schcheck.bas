Attribute VB_Name = "M�dulo1"
Sub ResaltarTextoSeleccionado()
    Dim rng As Range
    Dim match As Object
    Dim matches As Object
    
    ' Verificar si hay algo seleccionado en el documento
    If Selection.Type = wdSelectionIP Or Selection.Type = wdSelectionNormal Then
        ' Obtener el rango seleccionado
        Set rng = Selection.Range
    Else
        MsgBox "No hay texto seleccionado."
        Exit Sub
    End If
    
    ' Definir la expresi�n regular
    Set match = CreateObject("VBScript.RegExp")
    match.Pattern = "There are not (.*?) security headers"
    match.Global = True ' Buscar todas las ocurrencias en una l�nea
    
    ' Buscar todas las ocurrencias del patr�n en el texto seleccionado
    Set matches = match.Execute(rng.Text)
    
    ' Resaltar el texto entre las palabras "There are not" y "security headers"
    For Each m In matches
        With rng.Duplicate
            .Start = .Start + m.FirstIndex
            .End = .Start + m.Length
            .Font.Bold = True
            .Font.Color = wdColorRed
        End With
    Next m
End Sub

