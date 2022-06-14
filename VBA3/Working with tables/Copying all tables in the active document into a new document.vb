Sub CopyTablesToNewDoc() 
 Dim docOld As Document 
 Dim rngDoc As Range 
 Dim tblDoc As Table 
 
 If ActiveDocument.Tables.Count >= 1 Then 
  Set docOld = ActiveDocument 
  Set rngDoc = Documents.Add.Range(Start:=0, End:=0) 
  For Each tblDoc In docOld.Tables 
   tblDoc.Range.Copy 
   With rngDoc 
    .Paste 
    .Collapse Direction:=wdCollapseEnd 
    .InsertParagraphAfter 
    .Collapse Direction:=wdCollapseEnd 
   End With 
  Next 
 End If 
End Sub