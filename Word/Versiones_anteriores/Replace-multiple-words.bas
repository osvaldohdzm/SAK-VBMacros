Attribute VB_Name = "Módulo2"
Sub ReplaceWordsInDocument()
    Dim replace_words_dict As Object
    Set replace_words_dict = CreateObject("Scripting.Dictionary")

    ' Add your word replacement pairs to the dictionary
    replace_words_dict.Add "##-##-2023", "22-12-2023"
    replace_words_dict.Add "XXX-XXX-2023", "SHFDARM/DSIC/077/2023"
    replace_words_dict.Add "##-mes-2023", "22-diciembre-2023"
    replace_words_dict.Add "##/##/2023", "22/12/2023"
    ' Add more pairs as needed

    ' Replace words in the main document body
    ReplaceInSelection ActiveDocument.Content, replace_words_dict

    ' Replace words in the first page header
    ReplaceInHeader ActiveDocument.Sections(1).Headers(1).Range, replace_words_dict

    ' Replace words in the primary header
    ReplaceInHeader ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary).Range, replace_words_dict
End Sub

Sub ReplaceInSelection(rng As Object, replaceWords As Object)
    Dim word_to_find As Variant
    Dim replace_word As Variant

    For Each word_to_find In replaceWords.keys
        replace_word = replaceWords(word_to_find)

        With rng.Find
            .Text = word_to_find
            .Replacement.Text = replace_word
            .Wrap = 1 ' wdFindContinue
            .Execute Replace:=2 ' wdReplaceAll
        End With
    Next word_to_find
End Sub

Sub ReplaceInHeader(rng As Object, replaceWords As Object)
    Dim word_to_find As Variant
    Dim replace_word As Variant

    For Each word_to_find In replaceWords.keys
        replace_word = replaceWords(word_to_find)

        With rng.Find
            .Text = word_to_find
            .Replacement.Text = replace_word
            .Wrap = 1 ' wdFindContinue
            .Execute Replace:=2 ' wdReplaceAll
        End With
    Next word_to_find
End Sub

