Set Range1 = Selection.Words(1) 
Set Range2 = ActiveDocument.Words(3) 
If Range1.IsEqual(Range:=Range2) = True Then 
 Range1.Delete 
End If