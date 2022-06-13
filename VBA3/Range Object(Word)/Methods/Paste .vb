If ActiveDocument.Tables.Count >= 1 Then 
 ActiveDocument.Tables(1).Range.Copy 
 Documents.Add.Content.Paste 
End If

If Selection.Type <> wdSelectionIP Then 
 Selection.Copy 
 Set Range2 = ActiveDocument.Content 
 Range2.Collapse Direction:=wdCollapseEnd 
 Range2.Paste 
End If