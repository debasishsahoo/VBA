pos = Selection.EndKey(Unit:=wdLine, Extend:=wdMove)

If Selection.Information(wdWithInTable) = True Then 
 Selection.HomeKey Unit:=wdColumn, Extend:=wdMove 
 Selection.EndKey Unit:=wdColumn, Extend:=wdExtend 
End If

Selection.EndKey Unit:=wdStory, Extend:=wdMove