Attribute VB_Name = "Module1"
Sub ReemplazarCamposDesdeArchivo()

    Dim seleccionArchivo As FileDialog
    Dim rutaArchivo As String
    Dim textoArchivo As String
    Dim lineas As Variant
    Dim campo As String
    Dim valor As String
    Dim documento As Document
    
    ' Abrir el explorador de archivos para seleccionar el archivo de texto
    Set seleccionArchivo = Application.FileDialog(msoFileDialogFilePicker)
    
    With seleccionArchivo
        .Title = "Seleccionar archivo de texto con campos"
        .Filters.Clear
        .Filters.Add "Archivos de texto", "*.txt"
        
        If .Show = -1 Then ' Si el usuario selecciona un archivo y hace clic en Abrir
            rutaArchivo = .SelectedItems(1)
        Else
            Exit Sub ' Si el usuario cancela, salir del macro
        End If
    End With
    
    ' Leer el contenido del archivo seleccionado
    Open rutaArchivo For Input As #1
    textoArchivo = Input$(LOF(1), #1)
    Close #1
    
    ' Dividir el texto en líneas
    lineas = Split(textoArchivo, vbCrLf)
    
    ' Obtener el documento actual de Word
    Set documento = ThisDocument
    
    ' Recorrer cada línea y realizar los reemplazos
    For Each linea In lineas
        If InStr(linea, ":") > 0 Then
            campo = Trim(Left(linea, InStr(linea, ":") - 1))
            valor = Trim(Mid(linea, InStr(linea, ":") + 1))
            
            ' Realizar el reemplazo en el documento
            With documento.Content.Find
                .Text = campo
                .Replacement.Text = valor
                .Wrap = wdFindContinue
                .Execute Replace:=wdReplaceAll
            End With
        End If
    Next linea
    
    ' Informar al usuario que se han realizado los reemplazos
    MsgBox "Se han realizado los reemplazos correctamente.", vbInformation
    
End Sub

