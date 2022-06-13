Selection.Collapse Direction:=wdCollapseEnd 
If Selection.IsEndOfRowMark = True Then 
 Selection.Rows(1).Select 
End If