Attribute VB_Name = "M�dulo3"
Sub CentrarTablasHorizontalmente()
    Dim tbl As Table

    For Each tbl In ActiveDocument.Tables
        tbl.Rows.Alignment = wdAlignRowCenter
    Next tbl
End Sub

