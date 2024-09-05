Attribute VB_Name = "MÛdulo1"
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

