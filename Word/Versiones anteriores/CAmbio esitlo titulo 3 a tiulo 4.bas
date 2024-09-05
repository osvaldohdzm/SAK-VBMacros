Attribute VB_Name = "Módulo1"
Sub CambiarTitulo4aTitulo3()
    Dim para As Paragraph
    For Each para In ActiveDocument.Paragraphs
        If para.Style = "Título 3" Then
            para.Style = "Título 4"
        End If
    Next para
End Sub

