Sub Macro4()

  Selection.Find.ClearFormatting

  Selection.Find.Replacement.ClearFormatting

  With Selection.Find

    .Text = ""

    .Replacement.Text = ""

    .Forward = True

    .Wrap = wdFindContinue

    .Format = True

    .MatchCase = False

    .MatchWholeWord = False

    .MatchWildcards = False

    .MatchSoundsLike = False

    .MatchAllWordForms = False

  End With

  Selection.Find.Execute Replace:=wdReplaceAll

End Sub


Sub Macro4A()

  With Selection.Find

    .ClearFormatting

    .Replacement.ClearFormatting

    .Text = ""

    .Replacement.Text = ""

    .Font.Italic = True

    .Replacement.Font.Bold = True

    .Forward = True

    .Wrap = wdFindContinue

    .Format = True

    .MatchCase = False

    .MatchWholeWord = False

    .MatchWildcards = False

    .MatchSoundsLike = False

    .MatchAllWordForms = False

    .Execute Replace:=wdReplaceAll

  End With

End Sub