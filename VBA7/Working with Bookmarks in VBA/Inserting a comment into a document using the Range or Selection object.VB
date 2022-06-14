Sub CommentsCollectionObject()
   Dim MyText As String
   Dim MyRange As Object
   Set MyRange = ActiveDocument.Range
   MyText = "<Replace this with your text>"
  ' Selection Example:
  Selection.Comments.Add Range:=Selection.Range, Text:=MyText
  ' Range Example:
  MyRange.Comments.Add Range:=Selection.Range, Text:=MyText
End Sub