Attribute VB_Name = "M�dulo1"
Sub ReemplazarTextoCeldasSeleccionadasEnTodoElDocumento()
    Dim tbl As Table
    Dim celda As Cell
    Dim textoActual As String
    Dim textoNuevo As String
    Dim respuesta As VbMsgBoxResult
    
    ' Verificar si la selecci�n est� dentro de una tabla
    If Selection.Information(wdWithInTable) = False Then
        MsgBox "Por favor, selecciona celdas dentro de una tabla.", vbExclamation
        Exit Sub
    End If
    
    ' Obtener la tabla activa
    Set tbl = Selection.Tables(1)
    
    ' Recorrer cada celda seleccionada
    For Each celda In Selection.Cells
        ' Obtener el texto actual de la celda
        textoActual = celda.Range.Text
        ' Eliminar el car�cter final (marcador de fin de celda)
        textoActual = Left(textoActual, Len(textoActual) - 2)
        
        ' Mostrar el texto actual al usuario y solicitar el texto de reemplazo
        textoNuevo = InputBox("Texto actual: " & textoActual & vbCrLf & _
                              "Ingresa el nuevo texto para reemplazar en todo el documento (dejar vac�o para mantener el texto actual):", _
                              "Reemplazar texto en documento")
        
        ' Si el texto de reemplazo no se ingresa, mantener el texto actual
        If textoNuevo = "" Then
            textoNuevo = textoActual
        End If
        
        ' Confirmar si desea realizar el cambio
        respuesta = MsgBox("�Reemplazar todas las ocurrencias de '" & textoActual & "' con '" & textoNuevo & "' en el documento?", vbYesNo + vbQuestion)
        If respuesta = vbYes Then
            ' Usar la funci�n Find para reemplazar en todo el documento
            With ActiveDocument.Content.Find
                .Text = textoActual
                .Replacement.Text = textoNuevo
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
                .Execute Replace:=wdReplaceAll
            End With
        End If
    Next celda
    
    MsgBox "Operaci�n completada. Todas las ocurrencias han sido reemplazadas.", vbInformation
End Sub


