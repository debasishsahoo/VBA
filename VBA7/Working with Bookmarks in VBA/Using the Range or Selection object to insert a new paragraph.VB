Sub InsertParagraphMethod()
   Dim MyRange As Object
   Set MyRange = ActiveDocument.Range
   ' Selection Example:
   Selection.InsertParagraph
   ' Range Example:
   MyRange.Collapse Direction:=wdCollapseStart
   MyRange.InsertParagraph
End Sub