Set myRange = Selection.Range 
myRange.WholeStory 
myRange.Font.Name = "Arial"

If ActiveDocument.Comments.Count >= 1 Then 
 Set myRange = Activedocument.Comments(1).Range 
 myRange.WholeStory 
 myRange.Copy 
 Documents.Add.Content.Paste 
End If
