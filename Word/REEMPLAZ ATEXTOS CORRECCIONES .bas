Attribute VB_Name = "Módulo1"
Sub ReemplazosCodigoPlural_Final()
    ' Macro definitivo para reemplazar frases restantes en todo el documento
    Dim replacements As Variant
    Dim i As Long
    
    ' Reemplazos ordenados de más largos a más cortos para evitar solapamientos
    replacements = Array( _
        Array("durante la ejecución del análisis del proyecto", "durante la ejecución del análisis de los códigos de las aplicaciones evaluadas"), _
        Array("realizado en la aplicación", "realizado sobre los códigos de las aplicaciones evaluadas"), _
        Array("del código del proyecto", "de los códigos de las aplicaciones evaluadas"), _
        Array("del código del aplicativo", "de los códigos de las aplicaciones evaluadas"), _
        Array("el comportamiento del aplicativo", "el comportamiento de los códigos de las aplicaciones evaluadas"), _
        Array("del proyecto", "de los códigos de las aplicaciones evaluadas"), _
        Array("el proyecto", "los códigos de las aplicaciones evaluadas"), _
        Array("la aplicación", "los códigos de las aplicaciones evaluadas"), _
        Array("de la aplicación", "de los códigos de las aplicaciones evaluadas"), _
        Array("el código esté accesible para el análisis", "los códigos estén accesibles para el análisis"), _
        Array("en el código o binarios de los códigos de las aplicaciones evaluadas", "en los códigos y binarios de las aplicaciones evaluadas"), _
        Array("sobre los códigos de las aplicaciones evaluadas de Aplicaciones Vulnerabilidades 2S", "sobre los códigos de las aplicaciones evaluadas en el proyecto 'Aplicaciones Vulnerabilidades 2S'") _
    )
    
    ' Aplicar cada reemplazo en todo el documento
    For i = LBound(replacements) To UBound(replacements)
        With ActiveDocument.Content.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = replacements(i)(0)
            .Replacement.Text = replacements(i)(1)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .Execute Replace:=wdReplaceAll
        End With
    Next i
    
    MsgBox "Todos los reemplazos restantes se aplicaron correctamente.", vbInformation
End Sub


