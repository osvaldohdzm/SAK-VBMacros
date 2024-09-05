Attribute VB_Name = "M�dulo1"
Sub UTF8ReemplazarAcentos()
    Dim Hoja As Worksheet
    Dim Texto As String

    ' Cambia "Hoja1" por el nombre de la hoja en la que deseas realizar los reemplazos
    Set Hoja = ThisWorkbook.Sheets(1)

    ' Reemplaza los caracteres mal codificados
    With Hoja.UsedRange
        Texto = .Cells(1, 1).Text
        .Replace What:="ó", Replacement:="�", LookAt:=xlPart
        .Replace What:="�", Replacement:="�", LookAt:=xlPart
        .Replace What:="á", Replacement:="�", LookAt:=xlPart
        .Replace What:="ñ", Replacement:="�", LookAt:=xlPart
        .Replace What:="ú", Replacement:="�", LookAt:=xlPart
        .Replace What:="é", Replacement:="�", LookAt:=xlPart
        .Replace What:="ü", Replacement:="", LookAt:=xlPart
        .Replace What:="�", Replacement:="�", LookAt:=xlPart
        .Replace What:="�", Replacement:="", LookAt:=xlPart
        .Replace What:="�", Replacement:="�", LookAt:=xlPart
        .Replace What:="CR�TICO", Replacement:="CR�TICO", LookAt:=xlWhole
    End With
    

    MsgBox "Reemplazo de acentos completado.", vbInformation
End Sub

