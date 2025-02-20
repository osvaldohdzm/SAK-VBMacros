Attribute VB_Name = "Module1"
Sub FormatearComoTabla()

    ' Get a contiguous selection (optional, but recommended for robustness)
    Dim rngSelection As Range
    Set rngSelection = ActiveSheet.UsedRange ' Get the used range on the active sheet


    ' Check if any cells are selected
    If rngSelection Is Nothing Then
        MsgBox "Please select a range of cells to format as a table.", vbExclamation
        Exit Sub
    End If

    ' Create the table
    ActiveSheet.ListObjects.Add(xlSrcRange, rngSelection, , xlYes).Name = "Tabla1"

    ' Format the table (assuming you want header row and all cells)
    With ActiveSheet.ListObjects("Tabla1")
        .TableStyle = xlTableStyleLight1  ' Change table style as desired
        .Range.RowHeight = 15  ' Set row height for all rows
    End With

End Sub

