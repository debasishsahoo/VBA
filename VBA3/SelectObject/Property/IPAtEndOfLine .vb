Selection.Collapse Direction:=wdCollapseEnd 
If Selection.IPAtEndOfLine = False Then 
 Selection.EndKey Unit:=wdLine, Extend:=wdMove 
End If