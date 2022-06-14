Sub InsertAfterMethod()
  Dim MyText As String
  Dim MyRange As Object
  Set MyRange = ActiveDocument.Range
  MyText = "<Replace this with your text>"
  ' Selection Example:
  Selection.InsertAfter (MyText)
  ' Range Example:
  ' (Inserts text at the current position of the insertion point.)
  MyRange.Collapse
  MyRange.InsertAfter (MyText)
End Sub


Sub InsertBeforeMethod()
   Dim MyText As String
   Dim MyRange As Object
   Set MyRange = ActiveDocument.Range
   MyText = "<Replace this with your text>"
   ' Selection Example:
   Selection.InsertBefore (MyText)
   ' Range Example: Inserts text at the beginning
   ' of the active document.
   MyRange.InsertBefore (MyText)
End Sub
	