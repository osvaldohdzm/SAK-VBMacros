Attribute VB_Name = "M�dulo1"
Sub Lowercase()
 For Each Cell In Selection
        If Not Cell.HasFormula Then
            Cell.Value = LCase(Cell.Value)
        End If
    Next Cell
End Sub
