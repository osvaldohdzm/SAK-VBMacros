Attribute VB_Name = "WordGeneralMacros"
Sub ApplyFontAndBoldToStyles()
    ' Declara variables para el documento actual y los estilos
    Dim doc As Document
    Dim style As style

    ' Asigna el documento activo a la variable doc
    Set doc = ActiveDocument

    ' Recorre cada estilo en el documento
    For Each style In doc.Styles
        ' Aplica la fuente Arial tama�o 12 al estilo actual si es un estilo de p�rrafo o de car�cter
        If style.Type = wdStyleTypeParagraph Or style.Type = wdStyleTypeCharacter Then
            With style.Font
                .Name = "Arial"
                .Size = 12
            End With
        End If

        ' Aplica negrita si el nombre del estilo contiene "T�tulo" o "Titulo"
        If InStr(style.NameLocal, "T�tulo") > 0 Or InStr(style.NameLocal, "Titulo") > 0 Then
            style.Font.Bold = True
        End If
    Next style
End Sub



Sub ApplyArialToWholeDocument()
    ' Selecciona todo el contenido del documento
    Selection.WholeStory

    ' Aplica la fuente Arial a la selecci�n
    With Selection.Font
        .Name = "Arial"
    End With

    ' Limpia la selecci�n
    Selection.Collapse Direction:=wdCollapseStart
End Sub

