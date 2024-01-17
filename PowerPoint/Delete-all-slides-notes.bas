Attribute VB_Name = "M�dulo2"
Sub EliminarNotas()
    Dim sld As Slide
    
    ' Recorrer todas las diapositivas en la presentaci�n
    For Each sld In ActivePresentation.Slides
        ' Eliminar las notas de la diapositiva actual
        sld.NotesPage.Shapes.Range.Delete
    Next sld
    
    ' Mostrar mensaje de notificaci�n
    MsgBox "Se han eliminado todas las notas de las diapositivas.", vbInformation
End Sub

