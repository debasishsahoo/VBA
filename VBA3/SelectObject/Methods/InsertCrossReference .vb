With Selection 
 .Collapse Direction:=wdCollapseStart 
 .InsertBefore "For more information, see " 
 .Collapse Direction:=wdCollapseEnd 
 .InsertCrossReference ReferenceType:=wdRefTypeHeading, _ 
 ReferenceKind:=wdContentText, ReferenceItem:=1 
 .InsertAfter " on page " 
 .Collapse Direction:=wdCollapseEnd 
 .InsertCrossReference ReferenceType:=wdRefTypeHeading, _ 
 ReferenceKind:=wdPageNumber, ReferenceItem:=1 
 .InsertAfter "." 
End With