If ActiveDocument.Subdocuments.Count >= 1 Then 
 ActiveDocument.ActiveWindow.View.Type = wdMasterView 
 Selection.HomeKey unit:=wdStory, Extend:=wdMove 
 Selection.NextSubdocument 
End If