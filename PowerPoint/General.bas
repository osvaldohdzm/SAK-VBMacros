Attribute VB_Name = "Módulo1"
Sub ShowShapeDimensions()
    Dim shp As Shape

    ' Check if there is a selected shape
    If Not ActiveWindow.Selection Is Nothing Then
        ' Ensure that only one shape is selected
        If ActiveWindow.Selection.ShapeRange.Count = 1 Then
            Set shp = ActiveWindow.Selection.ShapeRange(1)
            
            ' Display the dimensions of the selected shape (width and height)
            MsgBox "Shape Type: " & shp.Type & vbCrLf & _
                   "Width: " & shp.Width & " points" & vbCrLf & _
                   "Height: " & shp.Height & " points", _
                   vbInformation, "Shape Dimensions"
        Else
            MsgBox "Please select exactly one shape.", vbExclamation, "Error"
        End If
    Else
        MsgBox "No shape is selected.", vbExclamation, "Error"
    End If
End Sub


Sub UngroupAllGroups()
    Dim sld As Slide
    Dim shp As Shape
    Dim groupItem As Shape
    
    ' Loop through each slide in the presentation
    For Each sld In ActivePresentation.Slides
        ' Loop through each shape in the slide
        For Each shp In sld.Shapes
            ' Check if the shape is a group
            If shp.Type = msoGroup Then
                ' Ungroup the group
                shp.Ungroup
            End If
        Next shp
    Next sld
    
    MsgBox "All groups have been ungrouped.", vbInformation, "Ungroup Complete"
End Sub

Sub DeleteShapesByDimensionsEqual()
    Dim sld As Slide
    Dim shp As Shape
    Dim userWidth As Single
    Dim userHeight As Single
    Dim shapeWidth As Single
    Dim shapeHeight As Single
    
    ' Prompt the user to input the width and height in points
    userWidth = InputBox("Enter the width of the shapes to delete (in points):", "Shape Width")
    userHeight = InputBox("Enter the height of the shapes to delete (in points):", "Shape Height")
    
    ' Loop through each slide in the presentation
    For Each sld In ActivePresentation.Slides
        ' Loop through each shape in the slide
        For Each shp In sld.Shapes
            ' Get the dimensions of the shape
            shapeWidth = shp.Width
            shapeHeight = shp.Height
            
            ' Check if the width and height match the user input
            If shapeWidth = userWidth And shapeHeight = userHeight Then
                ' Delete the shape if it matches the provided dimensions
                shp.Delete
            End If
        Next shp
    Next sld
    
    MsgBox "All matching shapes have been deleted.", vbInformation, "Shapes Deleted"
End Sub



Sub DeleteShapesByDimensionsLess()
    Dim sld As Slide
    Dim shp As Shape
    Dim userWidth As Single
    Dim userHeight As Single
    Dim shapeWidth As Single
    Dim shapeHeight As Single
    
    ' Prompt the user to input the width and height in points
    userWidth = InputBox("Enter the width of the shapes to delete (in points):", "Shape Width")
    userHeight = InputBox("Enter the height of the shapes to delete (in points):", "Shape Height")
    
    ' Loop through each slide in the presentation
    For Each sld In ActivePresentation.Slides
        ' Loop through each shape in the slide
        For Each shp In sld.Shapes
            ' Get the dimensions of the shape
            shapeWidth = shp.Width
            shapeHeight = shp.Height
            
            ' Check if the shape's width and height are less than the user input
            If shapeWidth < userWidth And shapeHeight < userHeight Then
                ' Delete the shape if both its width and height are smaller than the provided values
                shp.Delete
            End If
        Next shp
    Next sld
    
    MsgBox "All matching shapes have been deleted.", vbInformation, "Shapes Deleted"
End Sub

