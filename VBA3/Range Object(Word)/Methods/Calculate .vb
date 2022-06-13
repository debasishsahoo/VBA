Set myRange = ActiveDocument.Range(0, 0) 
myRange.InsertBefore "1 + 1 " 
myRange.InsertAfter "= " & myRange.Calculate