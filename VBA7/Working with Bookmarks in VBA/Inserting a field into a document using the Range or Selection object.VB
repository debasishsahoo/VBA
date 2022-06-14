Sub FieldsCollectionObject()
   Dim MyText As String
   Dim MyRange As Object
   Set MyRange = Selection.Range
   MyText = "<Replace this with your text>"
   ' Selection Example:
   Selection.Fields.Add Range:=Selection.Range, _
      Type:=wdFieldQuote, Text:=MyText
   ' Range Example:
    Range.Fields.Add Range:=Selection.Range, _
      Type:=wdFieldQuote, Text:=MyText
End Sub