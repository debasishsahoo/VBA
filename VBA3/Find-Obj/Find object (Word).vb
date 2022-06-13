With Selection.Find 
 .ClearFormatting 
 .Text = "hi" 
 .Execute Forward:=True 
End With


Set myRange = ActiveDocument.Content 
myRange.Find.Execute FindText:="hi", ReplaceWith:="hello", _ 
 Replace:=wdReplaceAll

Selection.Find.Execute FindText:="blue", Forward:=True

Set myRange = ActiveDocument.Content 
myRange.Find.Execute FindText:="blue", Forward:=True 
If myRange.Find.Found = True Then myRange.Bold = True