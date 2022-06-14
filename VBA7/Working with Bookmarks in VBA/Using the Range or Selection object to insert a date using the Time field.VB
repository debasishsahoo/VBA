Sub InsertDateTimeMethod()
   Dim MyRange As Object
   Set MyRange = Selection.Range
   ' Selection Example:
   Selection.InsertDateTime DateTimeFormat:="MMMM dd, yyyy", _
   InsertAsField:=True
   ' Range Example:
   MyRange.InsertDateTime DateTimeFormat:="MMM dd, yyyy", _
   InsertAsField:=True
End Sub