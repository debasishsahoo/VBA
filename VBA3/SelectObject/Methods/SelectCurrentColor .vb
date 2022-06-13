Selection.HomeKey Unit:=wdStory, Extend:=wdMove 
Selection.SelectCurrentColor 
n = Len(Selection.Text) 
MsgBox "Contiguous characters with the same color: " & n