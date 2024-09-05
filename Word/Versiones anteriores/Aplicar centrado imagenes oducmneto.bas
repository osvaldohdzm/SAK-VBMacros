Attribute VB_Name = "NewMacros"
Sub CentrarImagenesInline()
    Dim img As inlineShape
    For Each img In ActiveDocument.InlineShapes
        img.Select
        Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Next img
End Sub

