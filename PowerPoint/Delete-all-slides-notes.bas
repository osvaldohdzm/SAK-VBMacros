Attribute VB_Name = "Módulo2"
Sub EliminarNotas()
    Dim sld As Slide
    
    ' Recorrer todas las diapositivas en la presentación
    For Each sld In ActivePresentation.Slides
        ' Eliminar las notas de la diapositiva actual
        sld.NotesPage.Shapes.Range.Delete
    Next sld
    
    ' Mostrar mensaje de notificación
    MsgBox "Se han eliminado todas las notas de las diapositivas.", vbInformation
End Sub

