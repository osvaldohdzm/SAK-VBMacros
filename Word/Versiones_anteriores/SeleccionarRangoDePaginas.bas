Attribute VB_Name = "NewMacros"
Sub SeleccionarRangoDePaginas()
    Dim rangoPaginas As String
    Dim paginaInicio As Integer, paginaFin As Integer
    Dim inicioTexto As Range, finTexto As Range
    Dim paginasDocumento As Long
    
    ' Obtener el número total de páginas en el documento
    paginasDocumento = ActiveDocument.ComputeStatistics(wdStatisticPages)
    
    ' Solicitar al usuario el rango de páginas
    rangoPaginas = InputBox("Ingrese el rango de páginas en el formato 'inicio-fin' (por ejemplo, '2-5'):", "Seleccionar rango de páginas")
    
    ' Verificar si el formato del rango de páginas es válido
    If InStr(rangoPaginas, "-") > 0 Then
        ' Dividir el rango de páginas en página de inicio y página final
        paginaInicio = Val(Split(rangoPaginas, "-")(0))
        paginaFin = Val(Split(rangoPaginas, "-")(1))
        
        ' Verificar si las páginas están dentro del rango válido
        If paginaInicio > 0 And paginaFin > 0 And paginaInicio <= paginasDocumento And paginaFin <= paginasDocumento Then
            ' Ir al principio de la página de inicio
            Set inicioTexto = ActiveDocument.GoTo(What:=wdGoToPage, Name:=paginaInicio).GoTo(What:=wdGoToBookmark, Name:="\page")
            
            ' Ir al final de la página de fin
            Set finTexto = ActiveDocument.GoTo(What:=wdGoToPage, Name:=paginaFin).GoTo(What:=wdGoToBookmark, Name:="\page")
            
            ' Seleccionar todo el texto entre las páginas especificadas
            If inicioTexto.Start < finTexto.Start Then
                ActiveDocument.Range(inicioTexto.Start, finTexto.End).Select
            Else
                MsgBox "Error: La página de inicio debe ser anterior a la página de fin.", vbExclamation
            End If
        Else
            MsgBox "El rango de páginas especificado está fuera del rango válido.", vbExclamation
        End If
    Else
        MsgBox "El formato del rango de páginas es incorrecto. Por favor, ingrese un rango válido en el formato 'inicio-fin'.", vbExclamation
    End If
End Sub

