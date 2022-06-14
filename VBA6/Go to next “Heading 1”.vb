Sub GoToNextHeading1()
  Application.ScreenUpdating = False
  ' First, moves to end of Heading 1 range if insertion point is in H1
  While Selection.Characters.Last.Next.Style = "Heading 1"
    Selection.MoveEnd Unit:=wdCharacter, Count:=1
  Wend
  Selection.MoveStart Unit:=wdCharacter, Count:=1
  Selection.Collapse Direction:=wdCollapseEnd
  With Selection.Find
    .ClearFormatting
    .Text = ""
    .MatchWildcards = False
    .Forward = True
    .Style = ActiveDocument.Styles("Heading 1")
    .Execute
    .ClearFormatting
  End With
  Selection.Collapse Direction:=wdCollapseStart
  Application.ScreenUpdating = True
End Sub