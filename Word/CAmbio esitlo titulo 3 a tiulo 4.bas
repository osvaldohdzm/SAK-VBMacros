Attribute VB_Name = "M�dulo1"
Sub CambiarTitulo4aTitulo3()
    Dim para As Paragraph
    For Each para In ActiveDocument.Paragraphs
        If para.Style = "T�tulo 3" Then
            para.Style = "T�tulo 4"
        End If
    Next para
End Sub

