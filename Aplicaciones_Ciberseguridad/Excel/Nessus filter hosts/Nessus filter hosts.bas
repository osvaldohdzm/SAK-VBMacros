Attribute VB_Name = "Módulo1"
Sub ConcatenateIPs()
    Dim cell As Range
    Dim selectedRange As Range
    Dim result As String
    Dim counter As Integer
    Dim txtBox As OLEObject
    Dim closeButton As OLEObject

    ' Set the selected range
    Set selectedRange = Selection
    
    ' Initialize the result string and counter
    result = "GET /scans/1680?limit=2500&"
    counter = 0
    
    ' Loop through each cell in the selected range
    For Each cell In selectedRange
        If Not IsEmpty(cell.Value) Then
            result = result & "filter." & counter & ".quality=eq&filter." & counter & ".filter=hostname&filter." & counter & ".value=" & cell.Value & "&"
            counter = counter + 1
        End If
    Next cell
    
    ' Add the remaining part of the URL
    result = result & "filter.search_type=or&includeHostDetailsForHostDiscovery=true HTTP/1.1"
    
    ' Create and configure the TextBox
    On Error Resume Next
    ' Check if a TextBox already exists and delete it
    Set txtBox = ActiveSheet.OLEObjects("ConcatenatedTextBox")
    If Not txtBox Is Nothing Then
        txtBox.Delete
    End If
    On Error GoTo 0
    
    ' Create a new TextBox
    Set txtBox = ActiveSheet.OLEObjects.Add(ClassType:="Forms.TextBox.1", Link:=False, DisplayAsIcon:=False, _
                                            Left:=10, Top:=10, Width:=800, Height:=100)
    txtBox.Name = "ConcatenatedTextBox"
    
    With txtBox.Object
        .Text = result
        .MultiLine = True
        .EnterKeyBehavior = True
        .WordWrap = False
        .ScrollBars = fmScrollBarsBoth
    End With
    
    ' Create and configure the Close Button
    On Error Resume Next
    ' Check if a Close Button already exists and delete it
    Set closeButton = ActiveSheet.OLEObjects("CloseButton")
    If Not closeButton Is Nothing Then
        closeButton.Delete
    End If
    On Error GoTo 0
    
    ' Create a new Close Button
    Set closeButton = ActiveSheet.OLEObjects.Add(ClassType:="Forms.CommandButton.1", Link:=False, DisplayAsIcon:=False, _
                                                 Left:=10, Top:=120, Width:=100, Height:=30)
    closeButton.Name = "CloseButton"
    closeButton.Object.Caption = "Close"
    
    ' Assign the Close Button macro
    ActiveSheet.OLEObjects("CloseButton").Object.OnClick = "CloseTextBox"
End Sub

Sub CloseTextBox()
    ' Delete the TextBox and Close Button
    On Error Resume Next
    ActiveSheet.OLEObjects("ConcatenatedTextBox").Delete
    ActiveSheet.OLEObjects("CloseButton").Delete
    On Error GoTo 0
End Sub


