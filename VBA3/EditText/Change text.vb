Sub ChangeText() 
 ActiveDocument.Words(1).Text = "The " 
End Sub

Sub DeleteText() 
 Dim rngFirstParagraph As Range 
 
 Set rngFirstParagraph = ActiveDocument.Paragraphs(1).Range 
 With rngFirstParagraph 
 .Delete 
 .InsertAfter Text:="New text" 
 .InsertParagraphAfter 
 End With 
End Sub