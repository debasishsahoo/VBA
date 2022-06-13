If Options.ReplaceSelection = True Then 
 Selection.Collapse Direction:=wdCollapseStart 
 Selection.TypeText Text:="Hello" 
End If


Options.ReplaceSelection = False 
With Selection 
 .TypeText Text:="Title" 
 .TypeParagraph 
End With