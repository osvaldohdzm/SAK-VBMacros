Attribute VB_Name = "M�dulo1"
Sub UTF8ReemplazarAcentos()
    Dim Hoja As Worksheet
    Dim Texto As String

    ' Cambia "Hoja1" por el nombre de la hoja en la que deseas realizar los reemplazos
    Set Hoja = ThisWorkbook.Sheets(1)

    ' Reemplaza los caracteres mal codificados
    With Hoja.UsedRange
        Texto = .Cells(1, 1).Text
        .Replace What:="Ã³", Replacement:="ó", LookAt:=xlPart
        .Replace What:="í“", Replacement:="ó", LookAt:=xlPart
        .Replace What:="Ã¡", Replacement:="á", LookAt:=xlPart
        .Replace What:="Ã±", Replacement:="ñ", LookAt:=xlPart
        .Replace What:="Ãº", Replacement:="ú", LookAt:=xlPart
        .Replace What:="Ã©", Replacement:="é", LookAt:=xlPart
        .Replace What:="Ã¼", Replacement:="", LookAt:=xlPart
        .Replace What:="Ã", Replacement:="í", LookAt:=xlPart
        .Replace What:="Â", Replacement:="", LookAt:=xlPart
        .Replace What:="í­­­­", Replacement:="í", LookAt:=xlPart
        .Replace What:="â€”", Replacement:="", LookAt:=xlPart
        .Replace What:="€”", Replacement:="", LookAt:=xlPart
        .Replace What:="í­a", Replacement:="í­a", LookAt:=xlPart
        .Replace What:="CRíTICO", Replacement:="CRÍTICO", LookAt:=xlWhole
        .Replace What:="CRíTICA", Replacement:="CRÍTICA", LookAt:=xlWhole
        
    End With
    

    MsgBox "Reemplazo de acentos completado.", vbInformation
End Sub



