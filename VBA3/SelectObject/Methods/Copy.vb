If Selection.Type = wdSelectionNormal Then 
 Selection.Copy 
 Documents.Add.Content.Paste 
End If