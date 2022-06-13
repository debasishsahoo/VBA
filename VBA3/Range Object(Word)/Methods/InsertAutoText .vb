Documents.Add 
Selection.TypeText "Best w" 
Selection.Range.InsertAutoText

Documents.Add 
Selection.TypeText "In " 
Set myRange = ActiveDocument.Words(1) 
myRange.InsertAutoText
