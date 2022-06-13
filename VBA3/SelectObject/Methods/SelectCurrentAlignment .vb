Selection.SelectCurrentAlignment 
Selection.Collapse Direction:=wdCollapseEnd 
If Selection.End = ActiveDocument.Content.End - 1 Then 
 MsgBox "No change in alignment found." 
End If