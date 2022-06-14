Sub SelectToBracketsInclusive()
  With Selection.Find
    .ClearFormatting
    .Text = "["
    .Forward = False
    .MatchWildcards = False
    .Wrap = wdFindStop
    .Execute
  End With
  Selection.Extend
  With Selection.Find
    .Text = "]"
    .Forward = True
    .Execute
    .Text = ""
  End With
End Sub