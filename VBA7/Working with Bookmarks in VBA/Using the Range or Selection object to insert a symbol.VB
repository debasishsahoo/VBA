Sub InsertSymbolMethod()
   Dim MyRange As Object
   Set MyRange = ActiveDocument.Range
   ' Selection Example:
   Selection.InsertSymbol CharacterNumber:=171, _
      Font:="Symbol", Unicode:=False
   ' Range Example:
   MyRange.Collapse Direction:=wdCollapseStart
   MyRange.InsertSymbol CharacterNumber:=171, _
      Font:="Symbol", Unicode:=False
End Sub
	