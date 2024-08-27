Attribute VB_Name = "ExcelVBA"
Sub ConvertSelectionToTextFormat()
    Dim cell As Range
    
    ' Desactivar la actualizaciÛn de la pantalla para mejorar la eficiencia
    Application.ScreenUpdating = False
    
    ' Iterar a travÈs de cada celda en la selecciÛn
    For Each cell In Selection
        If IsDate(cell.Value) Then
            ' Convertir la fecha a texto en el formato deseado y almacenar en una variable
            Dim formattedDate As String
            formattedDate = Format(cell.Value, "yyyy-mm-dd")
            
            ' Establecer el formato de la celda a texto
            cell.NumberFormat = "@"
            
            ' Asignar el valor formateado a la celda
            cell.Value = formattedDate
        End If
    Next cell
    
    ' Reactivar la actualizaciÛn de la pantalla
    Application.ScreenUpdating = True
    
    MsgBox "ConversiÛn a formato de texto completada."
End Sub

Sub UTF8ReemplazarAcentos()
    Dim Hoja As Worksheet
    Dim Texto As String

    ' Cambia "Hoja1" por el nombre de la hoja en la que deseas realizar los reemplazos
    Set Hoja = ThisWorkbook.Sheets(1)

    ' Reemplaza los caracteres mal codificados
    With Hoja.UsedRange
        Texto = .Cells(1, 1).Text
        .Replace What:="√≥", Replacement:="Û", LookAt:=xlPart
        .Replace What:="Ìì", Replacement:="Û", LookAt:=xlPart
        .Replace What:="√°", Replacement:="·", LookAt:=xlPart
        .Replace What:="√±", Replacement:="Ò", LookAt:=xlPart
        .Replace What:="√∫", Replacement:="˙", LookAt:=xlPart
        .Replace What:="√©", Replacement:="È", LookAt:=xlPart
        .Replace What:="√º", Replacement:="", LookAt:=xlPart
        .Replace What:="√", Replacement:="Ì", LookAt:=xlPart
        .Replace What:="¬", Replacement:="", LookAt:=xlPart
        .Replace What:="Ì≠", Replacement:="Ì", LookAt:=xlPart
        .Replace What:="CRÌçTICO", Replacement:="CRÕTICO", LookAt:=xlWhole
    End With
    

    MsgBox "Reemplazo de acentos completado.", vbInformation
End Sub


