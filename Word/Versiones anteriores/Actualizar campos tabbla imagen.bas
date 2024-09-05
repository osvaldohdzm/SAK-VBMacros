Attribute VB_Name = "Módulo4"
Sub ActualizarCamposSEQ()
    Dim doc As Document
    Set doc = ActiveDocument
    
    ' Recorre todos los campos del documento
    For Each campo In doc.Fields
        ' Comprueba si el campo es de tipo SEQ
        If campo.Type = wdFieldSequence Then
            ' Actualiza el campo
            campo.Update
        End If
    Next campo
End Sub

