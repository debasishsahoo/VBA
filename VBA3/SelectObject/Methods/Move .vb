If Selection.StoryType = wdMainTextStory Then 
 wUnits = Selection.Move(Unit:=wdWord, Count:=2) 
 If wUnits < 2 Then _ 
 MsgBox "Selection is at the end of the document" 
End If


If Selection.Information(wdWithInTable) = True Then 
 Selection.Move Unit:=wdCell, Count:=3 
End If