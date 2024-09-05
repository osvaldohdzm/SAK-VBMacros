Attribute VB_Name = "NewMacros"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro1"
'
' Macro1 Macro
'
'
    Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
    With ListGalleries(wdNumberGallery).ListTemplates(1).ListLevels(1)
        .NumberFormat = "[%1]"
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(0.63)
        .TabPosition = wdUndefined
        .ResetOnHigher = 0
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdUndefined
            .Size = wdUndefined
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = ""
        End With
        .LinkedStyle = ""
    End With
    ListGalleries(wdNumberGallery).ListTemplates(1).Name = ""
    Selection.Range.ListFormat.ApplyListTemplateWithLevel ListTemplate:= _
        ListGalleries(wdNumberGallery).ListTemplates(1), ContinuePreviousList:= _
        True, ApplyTo:=wdListApplyToWholeList, DefaultListBehavior:= _
        wdWord10ListBehavior
End Sub
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro2"
'
' Macro2 Macro
'
'
    With ListGalleries(wdNumberGallery).ListTemplates(7).ListLevels(1)
        .NumberFormat = "[%1]"
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(0.63)
        .TabPosition = wdUndefined
        .ResetOnHigher = 0
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdUndefined
            .Size = wdUndefined
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = ""
        End With
        .LinkedStyle = "Referencia Tabla"
    End With
    ListGalleries(wdNumberGallery).ListTemplates(7).Name = ""
    Selection.Range.ListFormat.ApplyListTemplateWithLevel ListTemplate:= _
        ListGalleries(wdNumberGallery).ListTemplates(7), ContinuePreviousList:= _
        False, ApplyTo:=wdListApplyToWholeList, DefaultListBehavior:= _
        wdWord10ListBehavior
    With ListGalleries(wdNumberGallery).ListTemplates(7).ListLevels(1)
        .NumberFormat = "[%1]"
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(0.63)
        .TabPosition = wdUndefined
        .ResetOnHigher = 0
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdUndefined
            .Size = wdUndefined
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = ""
        End With
        .LinkedStyle = "Referencia Tabla"
    End With
    ListGalleries(wdNumberGallery).ListTemplates(7).Name = ""
    Selection.Range.ListFormat.ApplyListTemplateWithLevel ListTemplate:= _
        ListGalleries(wdNumberGallery).ListTemplates(7), ContinuePreviousList:= _
        False, ApplyTo:=wdListApplyToWholeList, DefaultListBehavior:= _
        wdWord10ListBehavior
End Sub
