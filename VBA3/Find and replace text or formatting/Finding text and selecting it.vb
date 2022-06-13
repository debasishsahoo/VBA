With Selection.Find 
 .Forward = True 
 .Wrap = wdFindStop 
 .Text = "Hello" 
 .Execute 
End With

Selection.Find.Execute FindText:="Hello", _ 
 Forward:=True, Wrap:=wdFindStop