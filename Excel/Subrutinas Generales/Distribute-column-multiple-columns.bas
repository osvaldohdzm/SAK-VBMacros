Attribute VB_Name = "M�dulo1"
Sub DistribuirCeldas()
    ' Obt�n la hoja actual
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet

    ' Obt�n las celdas seleccionadas
    Dim rng As Range
    Set rng = Selection

    ' Obt�n el n�mero de celdas seleccionadas
    Dim n As Long
    n = rng.Count

    ' Definir el n�mero de columnas como 4
    Dim numColumnas As Integer
    numColumnas = 4

    ' Calcular la longitud de cada columna
    Dim longitudColumna As Integer
    longitudColumna = WorksheetFunction.Ceiling(n / numColumnas, 1)

    ' Crea una nueva hoja
    Dim wsNew As Worksheet
    Set wsNew = ThisWorkbook.Sheets.Add(After:=ws)
    wsNew.Name = "DistributedColumns"

    ' Variables para realizar el bucle de distribuci�n
    Dim i As Integer
    Dim col As Integer
    Dim row As Integer

    ' Inicializa las variables de posici�n
    col = 1
    row = 1

    ' Itera sobre las celdas seleccionadas
    For i = 1 To n
        ' Escribe el valor de la celda en la nueva hoja
        wsNew.Cells(row, col).Value = rng.Cells(i).Value

        ' Mueve a la siguiente fila
        row = row + 1

        ' Verifica si es necesario pasar a la siguiente columna
        If row > longitudColumna Then
            col = col + 1
            row = 1
        End If
    Next i
End Sub

