Set myRange = ActiveDocument.Words(1) 
If myRange.InRange(Selection.Range) = False Then myRange.Select