If Selection.Type = wdSelectionNormal Then 
 Set Range1 = Selection.Range 
 Range1.TextRetrievalMode.IncludeHiddenText = False 
 Set Range2 = ActiveDocument.Paragraphs(2).Range 
 Range2.InsertAfter Range1.Text 
End If


Set myRange = ActiveDocument.Range(Start:=ActiveDocument _ 
 .Paragraphs(1).Range.Start, _ 
 End:=ActiveDocument.Paragraphs(3).Range.End) 
myRange.TextRetrievalMode.ViewType = wdOutlineView 
MsgBox myRange.Text


If Selection.Type = wdSelectionNormal Then 
 Set aRange = Selection.Range 
 With aRange.TextRetrievalMode 
 .IncludeHiddenText = False 
 .IncludeFieldCodes = False 
 End With 
 MsgBox aRange.Text 
End If