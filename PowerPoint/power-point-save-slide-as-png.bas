Attribute VB_Name = "M�dulo1"
Sub SaveCurrentSlideAsPNG()
    Dim sld As Slide
    Dim path As String
    Dim fileName As String
    Dim currentDate As String
    
    currentDate = Format(Now(), "yyyy-mm-dd-hh-nn")

    Set sld = ActiveWindow.View.Slide
    path = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"
    fileName = currentDate & ".png"
    sld.Export path & fileName, "PNG"
End Sub

