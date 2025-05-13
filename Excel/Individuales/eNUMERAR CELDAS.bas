Attribute VB_Name = "M�dulo3"
Sub EnumerarCeldas()
    Dim inicio As Long
    Dim celda As Range
    Dim seleccion As Range
    Dim valorActual As Long
    
    ' Solicitar al usuario el n�mero inicial
    On Error Resume Next
    inicio = Application.InputBox("Ingrese el n�mero inicial:", "Inicio de Enumeraci�n", Type:=1)
    On Error GoTo 0
    If inicio = 0 Or inicio = False Then Exit Sub ' Salir si se cancela o se ingresa un valor inv�lido
    
    valorActual = inicio ' Asignar el n�mero inicial
    
    ' Iterar sobre las celdas seleccionadas
    Set seleccion = Selection
    For Each celda In seleccion
        If Not celda.MergeCells Then ' Evitar celdas combinadas
            celda.Value = valorActual
            valorActual = valorActual + 1
        End If
    Next celda
    
    MsgBox "Enumeraci�n completada.", vbInformation
End Sub

