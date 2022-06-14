Sub SelectToBracketsExclusive()
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
  Selection.MoveStart Unit:=wdCharacter, Count:=1
  Selection.MoveEnd Unit:=wdCharacter, Count:=-1
End Sub