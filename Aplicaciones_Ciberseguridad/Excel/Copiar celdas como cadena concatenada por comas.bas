Attribute VB_Name = "Módulo1"
Attribute VB_Name = "Módulo1"
Sub ConcatenarConComas()
    Dim rng As Range
    Dim concatenado As String
    
    ' Verificar si hay celdas seleccionadas
    If Selection.Cells.Count > 0 Then
        ' Iterar a través de cada celda seleccionada
        For Each rng In Selection
            ' Concatenar el valor de la celda a la cadena con comas
            concatenado = concatenado & rng.Value & ","
        Next rng
        
        ' Eliminar la última coma extra
        concatenado = Left(concatenado, Len(concatenado) - 1)
        
        ' Copiar el texto concatenado al portapapeles
        CopyToClipboard concatenado
        
        ' Mostrar un mensaje informativo
        MsgBox "El texto concatenado se ha copiado al portapapeles.", vbInformation
    Else
        ' Si no hay celdas seleccionadas, mostrar un mensaje de error
        MsgBox "No hay celdas seleccionadas.", vbExclamation
    End If
End Sub

Sub CopyToClipboard(Text As String)
    ' VBA Macro using late binding to copy text to clipboard.
    Dim MSForms_DataObject As Object
    Set MSForms_DataObject = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    MSForms_DataObject.SetText Text
    MSForms_DataObject.PutInClipboard
    Set MSForms_DataObject = Nothing
End Sub


