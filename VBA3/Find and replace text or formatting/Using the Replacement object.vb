With Selection.Find 
 .ClearFormatting 
 .Text = "hi" 
 .Replacement.ClearFormatting 
 .Replacement.Text = "hello" 
 .Execute Replace:=wdReplaceAll, Forward:=True, _ 
 Wrap:=wdFindContinue 
End With

With ActiveDocument.Content.Find 
 .ClearFormatting 
 .Font.Bold = True 
 With .Replacement 
 .ClearFormatting 
 .Font.Bold = False 
 End With 
 .Execute FindText:="", ReplaceWith:="", _ 
 Format:=True, Replace:=wdReplaceAll 
End With