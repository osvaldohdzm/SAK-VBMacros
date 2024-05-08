Sub ConvertToLowercase()
    Dim rng As Range
    Set rng = Selection
    
    If Not rng Is Nothing Then
        Dim cell As Range
        For Each cell In rng
            cell.Value = LCase(cell.Value)
        Next cell
    End If
End Sub
