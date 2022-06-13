story = Selection.StoryType

ActiveDocument.ActiveWindow.View.Type = wdNormalView 
If Selection.StoryType = wdFootnotesStory Then _ 
 ActiveDocument.ActiveWindow.ActivePane.Close