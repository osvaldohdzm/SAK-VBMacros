Attribute VB_Name = "Módulo4"
Sub ApplyCustomShadowToAllPictures()
    Dim Slide As Slide
    Dim Shape As Shape
    
    For Each Slide In ActivePresentation.Slides
        For Each Shape In Slide.Shapes
            If Shape.Type = msoPicture Or Shape.Type = msoLinkedPicture Then
                With Shape
                    .Shadow.Visible = True
                    .Shadow.ForeColor.RGB = RGB(0, 0, 0) ' Black color
                    .Shadow.Size = 102
                     .Shadow.Blur = 16
      .Shadow.IncrementOffsetX 5
      .Shadow.IncrementOffsetY 5
      .Shadow.Transparency = 0.5
      
      .Line.Visible = msoTrue
      .Line.Weight = 1
      .Line.Transparency = 0.5
                End With
            End If
        Next Shape
    Next Slide
End Sub

