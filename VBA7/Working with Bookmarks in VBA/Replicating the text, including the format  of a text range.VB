Sub FormattedTextProperty()
   ' This example copies the first paragraph in the document, including
   ' its formatting, and inserts the formatted text at the insertion
   ' point.
   Selection.Collapse Direction:=wdCollapseStart
   Selection.FormattedText = ActiveDocument.Paragraphs(1).Range
End Sub