If Selection.Type = wdSelectionNormal Then 
 Selection.Cut 
 Documents.Add.Content.Paste 
End If