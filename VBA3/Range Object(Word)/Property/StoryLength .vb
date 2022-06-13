Set myRange = ActiveDocument.Sections(1) _ 
 .Headers(wdHeaderFooterPrimary).Range 
If myRange.StoryLength > 1 Then MsgBox myRange.Text


If ActiveDocument.Content.StoryLength = 1 Then _ 
 ActiveDocument.Close SaveChanges:=wdDoNotSaveChanges