Slist = Selection.Range.SynonymInfo.SynonymList(Meaning:=1) 
For i = 1 To UBound(Slist) 
 Msgbox Slist(i) 
Next i