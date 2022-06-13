ActiveDocument.Content.InsertAfter "end of document"

response = InputBox("Type some text") 
With ActiveDocument.Paragraphs(1).Range 
 .InsertAfter "1." & Chr(9) & response 
 .InsertParagraphAfter 
End With
