With Selection.Find 
 .Text = "hi" 
 .ClearFormatting 
 .Replacement.Text = "hello" 
 .Replacement.ClearFormatting 
 .Execute Replace:=wdReplaceOne, Forward:=True 
End With


With ActiveDocument.Content.Find 
 .ClearFormatting 
 .Font.Bold = True 
 .Text = "" 
 With .Replacement 
 .ClearFormatting 
 .Font.Bold = False 
 .Text = "" 
 End With 
 .Execute Format:=True, Replace:=wdReplaceAll 
End With