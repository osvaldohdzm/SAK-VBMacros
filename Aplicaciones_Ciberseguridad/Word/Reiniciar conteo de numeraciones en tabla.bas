Attribute VB_Name = "Módulo3"
Sub ReiniciarConteoEnumeracionesEnTablas()
    Dim tabla As Table
    Dim rngCell As Range
    Dim para As Paragraph
    
    ' Definir la subcadena deseada
    Const subcadena As String = "REFERENCIA"
    
    ' Definir el estilo de lista deseado
    Const estiloLista As String = "Enumeración 1"
    
    ' Iterar sobre todas las tablas en el documento
    For Each tabla In ActiveDocument.Tables
        ' Verificar si la tabla tiene al menos 4 filas y 2 columnas
        If tabla.Rows.Count >= 4 And tabla.Columns.Count >= 2 Then
            ' Verificar si la celda (4, 1) contiene la subcadena
            If ContieneSubcadena(tabla.cell(4, 1).Range.Text, subcadena) Then
                ' Obtener la celda (4, 2) de la tabla
                Set rngCell = Nothing
                On Error Resume Next
                Set rngCell = tabla.cell(4, 2).Range
                On Error GoTo 0
                
                ' Verificar si se encontró la celda
                If Not rngCell Is Nothing Then
                    ' Aplicar el formato deseado a la lista en la celda (4, 2)
                    With rngCell.ListFormat.ListTemplate.ListLevels(1)
                        .NumberFormat = "[%1] "
                        .TrailingCharacter = wdTrailingTab
                        .NumberStyle = wdListNumberStyleArabic
                        .NumberPosition = CentimetersToPoints(0.63)
                        .Alignment = wdListLevelAlignLeft
                        .TextPosition = CentimetersToPoints(1.27)
                        .TabPosition = wdUndefined
                        .ResetOnHigher = 0
                        .StartAt = 1
                        With .Font
                            .Bold = wdUndefined
                            .Italic = wdUndefined
                            .StrikeThrough = wdUndefined
                            .Subscript = wdUndefined
                            .Superscript = wdUndefined
                            .Shadow = wdUndefined
                            .Outline = wdUndefined
                            .Emboss = wdUndefined
                            .Engrave = wdUndefined
                            .AllCaps = wdUndefined
                            .Hidden = wdUndefined
                            .Underline = wdUndefined
                            .Color = wdUndefined
                            .Size = wdUndefined
                            .Animation = wdUndefined
                            .DoubleStrikeThrough = wdUndefined
                            .Name = ""
                        End With
                        
                    End With
                    ' Aplicar el estilo de lista
                    rngCell.ListFormat.ApplyListTemplateWithLevel ListTemplate:=ListGalleries(wdOutlineNumberGallery).ListTemplates(1), _
                    ContinuePreviousList:=False, ApplyTo:=wdListApplyToWholeList, DefaultListBehavior:=wdWord10ListBehavior
                End If
            End If
        End If
    Next tabla
End Sub

Function ContieneSubcadena(ByVal texto As String, ByVal subcadena As String) As Boolean
    ContieneSubcadena = InStr(1, texto, subcadena, vbTextCompare) > 0
End Function


