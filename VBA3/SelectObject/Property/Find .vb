With Selection.Find 
 .Forward = True 
 .ClearFormatting 
 .MatchWholeWord = True 
 .MatchCase = False 
 .Wrap = wdFindContinue 
 .Execute FindText:="Microsoft" 
End With