Sub SelectToBracketsDelete()
  With Selection.Find
    .ClearFormatting
    .Text = "["
    .Forward = False
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
  ' By not using "delete" method we get around Word's habit of adding spaces
  Selection.Text = ""
End Sub