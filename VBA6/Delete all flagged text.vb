Sub DeleteAllFlaggedText()
  ' First, confirm that user wants to do this
  Dim varResponse As Variant
  varResponse = MsgBox("Delete all flagged text?", vbYesNo, "Selection")
  If varResponse <> vbYes Then Exit Sub
  Dim oRng As Range
  Set oRng = ActiveDocument.Range(Start:=0, End:=0)
  With oRng.Find
    ' Preparation
    .ClearFormatting
    .Replacement.ClearFormatting
    .Text = ""
    .Replacement.Text = ""
    .MatchWildcards = False
    .Wrap = wdFindContinue
    ' Remove Flag 1 text
    .Style = ActiveDocument.Styles("Flag 1")
    .Execute Replace:=wdReplaceAll
    ' Remove Flag 2 text
    .Style = ActiveDocument.Styles("Flag 2")
    .Execute Replace:=wdReplaceAll
    ' Remove Flag 3 text
    .Style = ActiveDocument.Styles("Flag 3")
    .Execute Replace:=wdReplaceAll
    ' Clean out empty bracket pairs that once held flagged text
    .ClearFormatting
    .Text = "[]"
    .Execute Replace:=wdReplaceAll
    ' Clean up
    .Text = ""
    .Wrap = wdFindAsk
  End With
End Sub