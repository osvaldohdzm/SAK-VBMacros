Attribute VB_Name = "Módulo1"
Sub Uppercase()
  For Each Cell In Selection
    If Not Cell.HasFormula Then
      Cell.Value = UCase(Cell.Value)
    End If
  Next Cell
End Sub
