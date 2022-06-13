Sub InsertTextAtEndOfDocument() 
 ActiveDocument.Content.InsertAfter Text:=" The end." 
End Sub

Sub AddTextBeforeSelection() 
 Selection.InsertBefore Text:="new text " 
End Sub