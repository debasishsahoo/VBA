With ActiveDocument.Content.Find 
 .Text = "blue" 
 .Forward = True 
 .Execute 
 If .Found = True Then .Parent.Bold = True 
End With


Set myRange = ActiveDocument.Content 
myRange.Find.Execute FindText:="blue", Forward:=True 
If myRange.Find.Found = True Then myRange.Bold = True