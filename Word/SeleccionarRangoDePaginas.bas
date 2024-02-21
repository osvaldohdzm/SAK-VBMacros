Attribute VB_Name = "NewMacros"
Sub SeleccionarRangoDePaginas()
    Dim rangoPaginas As String
    Dim paginaInicio As Integer, paginaFin As Integer
    Dim inicioTexto As Range, finTexto As Range
    Dim paginasDocumento As Long
    
    ' Obtener el n�mero total de p�ginas en el documento
    paginasDocumento = ActiveDocument.ComputeStatistics(wdStatisticPages)
    
    ' Solicitar al usuario el rango de p�ginas
    rangoPaginas = InputBox("Ingrese el rango de p�ginas en el formato 'inicio-fin' (por ejemplo, '2-5'):", "Seleccionar rango de p�ginas")
    
    ' Verificar si el formato del rango de p�ginas es v�lido
    If InStr(rangoPaginas, "-") > 0 Then
        ' Dividir el rango de p�ginas en p�gina de inicio y p�gina final
        paginaInicio = Val(Split(rangoPaginas, "-")(0))
        paginaFin = Val(Split(rangoPaginas, "-")(1))
        
        ' Verificar si las p�ginas est�n dentro del rango v�lido
        If paginaInicio > 0 And paginaFin > 0 And paginaInicio <= paginasDocumento And paginaFin <= paginasDocumento Then
            ' Ir al principio de la p�gina de inicio
            Set inicioTexto = ActiveDocument.GoTo(What:=wdGoToPage, Name:=paginaInicio).GoTo(What:=wdGoToBookmark, Name:="\page")
            
            ' Ir al final de la p�gina de fin
            Set finTexto = ActiveDocument.GoTo(What:=wdGoToPage, Name:=paginaFin).GoTo(What:=wdGoToBookmark, Name:="\page")
            
            ' Seleccionar todo el texto entre las p�ginas especificadas
            If inicioTexto.Start < finTexto.Start Then
                ActiveDocument.Range(inicioTexto.Start, finTexto.End).Select
            Else
                MsgBox "Error: La p�gina de inicio debe ser anterior a la p�gina de fin.", vbExclamation
            End If
        Else
            MsgBox "El rango de p�ginas especificado est� fuera del rango v�lido.", vbExclamation
        End If
    Else
        MsgBox "El formato del rango de p�ginas es incorrecto. Por favor, ingrese un rango v�lido en el formato 'inicio-fin'.", vbExclamation
    End If
End Sub

