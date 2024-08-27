Attribute VB_Name = "FuncExcel"
Function ExtractTextBeforeSymbols(cell As Range) As String
    Dim text As String
    Dim posLessThan As Long
    Dim posParenthesis As Long
    Dim firstPos As Long
    
    text = cell.Value
    posLessThan = InStr(text, "<")
    posParenthesis = InStr(text, "(")
    
    ' Determine which position is the first one and greater than 0
    If posLessThan > 0 And posParenthesis > 0 Then
        firstPos = Application.WorksheetFunction.Min(posLessThan, posParenthesis)
    ElseIf posLessThan > 0 Then
        firstPos = posLessThan
    ElseIf posParenthesis > 0 Then
        firstPos = posParenthesis
    Else
        firstPos = 0
    End If
    
    If firstPos > 0 Then
        ExtractTextBeforeSymbols = Left(text, firstPos - 1)
    Else
        ExtractTextBeforeSymbols = text
    End If
End Function

