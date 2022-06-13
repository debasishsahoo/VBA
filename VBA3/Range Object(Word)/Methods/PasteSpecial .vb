Selection.Collapse Direction:=wdCollapseStart 
Selection.Range.PasteSpecial DataType:=wdPasteText

If Selection.Type = wdSelectionNormal Then 
 Selection.Copy 
 Documents.Add.Content.PasteSpecial Link:=True, _ 
 DataType:=wdPasteHyperlink 
End If
