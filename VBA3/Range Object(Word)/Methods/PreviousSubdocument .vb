If ActiveDocument.Subdocuments.Count >= 1 Then 
 ActiveDocument.ActiveWindow.View.Type = wdMasterView 
 Selection.EndKey Unit:=wdStory, Extend:=wdMove 
 Selection.PreviousSubdocument 
End If