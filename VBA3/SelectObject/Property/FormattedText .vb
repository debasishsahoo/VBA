Selection.Collapse Direction:=wdCollapseStart 
Selection.FormattedText = ActiveDocument.Paragraphs(1).Range

Set myRange = Selection.FormattedText 
Documents.Add.Content.FormattedText = myRange