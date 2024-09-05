Attribute VB_Name = "M�dulo11"
Sub EliminarSaltosDeLineaDentroDeParrafos()
    Dim Celda As Range
    Dim Texto As String
    Dim Linea As Variant
    Dim Parrafo As String
    Dim EnParrafo As Boolean
    Dim SaltosContiguos As Integer
    
    ' Itera a trav�s de las celdas seleccionadas en la hoja activa
    For Each Celda In Selection
        If Celda.HasFormula = False Then ' Ignora celdas con f�rmulas
            Texto = Celda.Value
            Parrafo = ""
            EnParrafo = False
            SaltosContiguos = 0
            
            ' Divide el texto en l�neas
            Dim Lineas As Variant
            Lineas = Split(Texto, vbLf) ' Utiliza vbLf para dividir por saltos de l�nea
            
            ' Recorre las l�neas y gestiona los saltos de l�nea
            For Each Linea In Lineas
                If Trim(Linea) = "" Then
                    ' Salto de l�nea en blanco, mant�nlo
                    Parrafo = Parrafo & vbCrLf
                Else
                    If SaltosContiguos > 1 Then
                        ' Mantener dos saltos contiguos
                        Parrafo = Parrafo & vbCrLf & vbCrLf & Linea
                    ElseIf SaltosContiguos = 1 Then
                        ' Mantener un solo salto
                        Parrafo = Parrafo & vbCrLf & " " & Linea
                    Else
                        ' No hay saltos contiguos, a�ade l�nea sin cambio
                        Parrafo = Parrafo & Linea
                    End If
                    SaltosContiguos = 0
                End If
                
                If Trim(Linea) = "" Then
                    SaltosContiguos = SaltosContiguos + 1
                End If
            Next Linea
            

              ' Elimina los espacios al inicio de cada l�nea
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

