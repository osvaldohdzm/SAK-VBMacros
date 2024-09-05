Attribute VB_Name = "NewMacros"

Sub Macro1()
    Dim corrections As Object
    Set corrections = CreateObject("Scripting.Dictionary")
    
    ' Agrega las correcciones al diccionario
    corrections("í­") = "í"
    corrections("e-") = "e"
    ' Añade tantas correcciones como necesites
    
    Dim key As Variant
    For Each key In corrections.Keys
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = key
            .Replacement.Text = corrections(key)
            .Forward = True
            .Wrap = wdFindAsk
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
    Next key
End Sub

