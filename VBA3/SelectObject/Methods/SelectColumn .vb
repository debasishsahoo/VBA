Selection.Collapse Direction:=wdCollapseEnd 
If Selection.Information(wdWithInTable) = True Then 
 Selection.SelectColumn 
End If