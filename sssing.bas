Sub AsignarNetMask()
    Dim ws As Worksheet
    Dim lr As Long
    Dim i As Long
    Dim dic As Object
    Dim v As Variant
    Dim colArray() As String
Dim col As Variant


colArray = Split(InputBox("Ingrese las letras de las columnas separadas por comas. Ejemplo: T,O,P,Q"), ",")

    
For Each col In colArray

    Set ws = ActiveSheet 'cambia esto por el nombre de tu hoja
    Set dic = CreateObject("Scripting.Dictionary")
    
    lr = ws.Cells(Rows.Count, "D").End(xlUp).Row 'asume que la columna D es PrivateIP
 


    For i = 2 To lr 'asume que la primera fila son los encabezados
        v = ws.Cells(i, "D").Value 'valor de PrivateIP
        If Not dic.Exists(v) Then 'si no está en el diccionario
            If Not IsEmpty(ws.Cells(i, col)) Then 'si la celda de NetMask no está vacía
                dic.Add v, ws.Cells(i, col).Value 'agrega el par PrivateIP-NetMask al diccionario
            End If
        End If
    Next i
    
    For i = 2 To lr 'recorre de nuevo la tabla
        v = ws.Cells(i, "D").Value 'valor de PrivateIP
        If dic.Exists(v) Then 'si está en el diccionario
            ws.Cells(i, col).Value = dic(v) 'asigna el valor de NetMask correspondiente
        End If
    Next i
    
    Set dic = Nothing

Next col

End Sub