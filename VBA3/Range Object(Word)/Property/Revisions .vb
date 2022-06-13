MsgBox ActiveDocument.Sections(1).Range.Revisions.Count

Set myRange = Selection.Paragraphs(1).Range 
myRange.Revisions.AcceptAll
