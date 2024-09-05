Attribute VB_Name = "M�dulo2"
Sub EliminarCodeSnippet()
    Dim rng As Range
    Dim doc As Document
    Set doc = ActiveDocument
    Set rng = doc.Content
    With rng.Find
        .ClearFormatting
        .Text = ""
        .Style = "CodeSnippet" ' Cambia "CodeSnippet" al nombre exacto de tu estilo de c�digo
        .Forward = True
        .Wrap = wdFindStop
        Do While .Execute
            rng.Select
            Selection.Delete
        Loop
    End With
End Sub

