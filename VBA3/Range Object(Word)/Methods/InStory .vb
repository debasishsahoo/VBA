Set Range1 = Selection.Words(1) 
Set Range2 = ActiveDocument.Range(Start:=20, End:=100) 
If Range1.InStory(Range:=Range2) = True Then 
 Range1.Font.Bold = True 
End If