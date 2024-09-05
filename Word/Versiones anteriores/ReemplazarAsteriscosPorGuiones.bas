Attribute VB_Name = "Módulo1"
Sub ReemplazarAsteriscosPorGuiones()
    ' Definir el objeto del rango del documento
    Dim Rng As Range
    ' Establecer el rango al contenido del documento
    Set Rng = ActiveDocument.Content
    ' Reemplazar asteriscos (*) por guiones (-)
    With Rng.Find
        .Text = "*"
        .Replacement.Text = "-"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With
End Sub

