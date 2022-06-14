Sub FindNextInstanceOfFlag1()
  Application.ScreenUpdating = False
  ' First, moves to end of Flag1 range if insertion point is in Flag1
  While Selection.Characters.Last.Next.Style = "Flag 1"
    Selection.MoveEnd Unit:=wdCharacter, Count:=1
  Wend
  Selection.Collapse Direction:=wdCollapseEnd
  With Selection.Find
  .ClearFormatting
  .Forward = True
  .Text = ""
  .MatchWildcards = False
  .Style = ActiveDocument.Styles("Flag 1")
  .Execute
  .ClearFormatting
  End With
  Selection.Collapse Direction:=wdCollapseStart
  Application.ScreenUpdating = True
End Sub