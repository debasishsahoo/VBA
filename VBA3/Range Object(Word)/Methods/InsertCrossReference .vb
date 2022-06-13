Set myRange = ActiveDocument.Range(Start:=0, End:=0) 
myBookmarks = ActiveDocument _ 
 .GetCrossReferenceItems(wdRefTypeBookmark) 
With myRange 
 .InsertBefore "Page " 
 .Collapse Direction:=wdCollapseEnd 
 .InsertCrossReference ReferenceType:=wdRefTypeBookmark, _ 
 ReferenceKind:=wdPageNumber, ReferenceItem:=myBookmarks(1) 
End With