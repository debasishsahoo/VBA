Sub SetStyleToFlag1()
  ' If no text selected, select text within the nearest brackets
  If Selection.Start = Selection.End Then SelectToBracketsExclusive
  Selection.Style = ActiveDocument.Styles("Flag 1")
End Sub