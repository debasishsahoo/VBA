With Selection 
 .HomeKey Unit:=wdStory, Extend:=wdMove 
 .SelectCurrentIndent 
 .Collapse Direction:=wdCollapseEnd 
End With

With Selection 
 .HomeKey Unit:=wdStory, Extend:=wdMove 
 .SelectCurrentIndent 
 .Collapse Direction:=wdCollapseEnd 
End With 
If Selection.End = ActiveDocument.Content.End - 1 Then 
 MsgBox "All paragraphs share the same left " _ 
 & "and right indents." 
Else 
 MsgBox "Not all paragraphs share the same left " _ 
 & "and right indents." 
End If