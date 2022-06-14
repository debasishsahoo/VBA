Sub PasteMethod()
   Dim MyRange As Object
   Set MyRange = Selection.Range
   ' Selection Example:
   Selection.Paste
   ' Range Example:
   MyRange.Collapse Direction:=wdCollapseStart
   MyRange.Paste
End Sub