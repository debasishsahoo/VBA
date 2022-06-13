status = Selection.InRange(ActiveDocument.Paragraphs(1).Range)

If Selection.InRange(ActiveDocument _ 
 .StoryRanges(wdFootnotesStory)) Then 
 MsgBox "Selection in footnotes" 
End If