Sub SelectParagraph() 
 ActiveDocument.Paragraphs(1).Range.Select 
 Selection.Font.Bold = True 
End Sub