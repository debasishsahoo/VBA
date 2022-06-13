Set myDoc = ActiveDocument 
Set myRange = _ 
 myDoc.Range(Start:= myDoc.Paragraphs(3).Range.Start, _ 
 End:=myDoc.Paragraphs(6).Range.End) 
myRange.ListFormat.ApplyOutlineNumberDefault


Selection.Range.ListFormat.ApplyListTemplate _ 
 ListTemplate:=ListGalleries(wdNumberGallery).ListTemplates(2)