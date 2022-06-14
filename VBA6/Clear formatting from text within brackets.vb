Sub ClearTextWithinBrackets()
  ' Removes char style / formatting within brackets, deletes the brackets
  Application.ScreenUpdating = False
  ActiveDocument.Bookmarks.Add Name:="LastPosition"
  ' First step: select text within brackets
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
  ' Second step: clear formatting in selected text and collapse selection
  With Selection
    .ClearFormatting
    .Collapse Direction:=wdCollapseStart
  End With
  ' Third step: delete the initial bracket
  With Selection
    .MoveEnd Unit:=wdCharacter, Count:=1
    .Text = ""
  End With
  ' Fourth step: find and delete the final bracket
  With Selection.Find
    .Text = "]"
    .Forward = True
    .Replacement.Text = ""
    .Execute Replace:=wdReplaceOne
    .Text = ""
  End With
  Selection.GoTo What:=wdGoToBookmark, Name:="LastPosition"
  ActiveDocument.Bookmarks(Index:="LastPosition").Delete
  Application.ScreenUpdating = True
End Sub