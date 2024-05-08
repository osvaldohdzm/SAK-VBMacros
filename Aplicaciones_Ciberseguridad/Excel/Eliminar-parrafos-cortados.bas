Attribute VB_Name = "Módulo11"
Sub EliminarSaltosDeLineaDentroDeParrafos()
    Dim Celda As Range
    Dim Texto As String
    Dim Linea As Variant
    Dim Parrafo As String
    Dim EnParrafo As Boolean
    Dim SaltosContiguos As Integer
    
    ' Itera a través de las celdas seleccionadas en la hoja activa
    For Each Celda In Selection
        If Celda.HasFormula = False Then ' Ignora celdas con fórmulas
            Texto = Celda.Value
            Parrafo = ""
            EnParrafo = False
            SaltosContiguos = 0
            
            ' Divide el texto en líneas
            Dim Lineas As Variant
            Lineas = Split(Texto, vbLf) ' Utiliza vbLf para dividir por saltos de línea
            
            ' Recorre las líneas y gestiona los saltos de línea
            For Each Linea In Lineas
                If Trim(Linea) = "" Then
                    ' Salto de línea en blanco, manténlo
                    Parrafo = Parrafo & vbCrLf
                Else
                    If SaltosContiguos > 1 Then
                        ' Mantener dos saltos contiguos
                        Parrafo = Parrafo & vbCrLf & vbCrLf & Linea
                    ElseIf SaltosContiguos = 1 Then
                        ' Mantener un solo salto
                        Parrafo = Parrafo & vbCrLf & " " & Linea
                    Else
                        ' No hay saltos contiguos, añade línea sin cambio
                        Parrafo = Parrafo & Linea
                    End If
                    SaltosContiguos = 0
                End If
                
                If Trim(Linea) = "" Then
                    SaltosContiguos = SaltosContiguos + 1
                End If
            Next Linea
            

              ' Elimina los espacios al inicio de cada línea
                Lineas = Split(Trim(Parrafo), vbCrLf)
                Parrafo = ""
                For i = LBound(Lineas) To UBound(Lineas)
                    Lineas(i) = Trim(Lineas(i))
                    If i = 0 Then
                        Parrafo = Lineas(i)
                    Else
                        Parrafo = Parrafo & vbCrLf & Lineas(i)
                    End If
                Next i
                
                ' Asigna el texto modificado a la celda
                Celda.Value = Parrafo
        End If
    Next Celda
End Sub

