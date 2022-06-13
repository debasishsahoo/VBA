Selection.HomeKey Unit:=wdStory, Extend:=wdMove

pos = Selection.HomeKey(Unit:=wdLine, Extend:=wdMove) 
If pos = 0 Then StatusBar = "Selection was not moved"