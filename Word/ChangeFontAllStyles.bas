Attribute VB_Name = "M�dulo1"
Sub CambiarFuenteAMontserrat()
    Dim estilo As Style
    
    For Each estilo In ActiveDocument.Styles
        If estilo.InUse Then
            estilo.Font.Name = "Montserrat"
        End If
    Next estilo
End Sub

