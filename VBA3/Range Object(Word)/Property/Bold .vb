If Selection.Type = wdSelectionNormal Then 
 Selection.Range.Bold = wdToggle 
End If

ActiveDocument.Paragraphs(1).Range.Bold = True
