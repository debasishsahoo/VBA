Sub GoToPreviousHeading1()
  Application.ScreenUpdating = False
  ' First, move to start of Heading 1 range if insertion point is in H1
  While Selection.Characters.First.Style = "Heading 1"
    Selection.MoveStart Unit:=wdCharacter, Count:=-1
  Wend
  Selection.Collapse Direction:=wdCollapseStart
  With Selection.Find
    .ClearFormatting
    .Text = ""
    .MatchWildcards = False
    .Forward = False
    .Style = ActiveDocument.Styles("Heading 1")
    .Execute
    .ClearFormatting
    .Forward = True
  End With
  Selection.Collapse Direction:=wdCollapseStart
  Application.ScreenUpdating = True
End Sub