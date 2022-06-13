ActiveDocument.Range.Collapse Direction:=wdCollapseEnd 
If ActiveDocument.Range.IsEndOfRowMark = True Then 
 ActiveDocument.Range.Rows(1).Select 
End If