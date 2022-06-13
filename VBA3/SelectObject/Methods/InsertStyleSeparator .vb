Sub InlineHeading() 
 Dim intCount As Integer 
 Dim intParaCount As Integer 
 
 intCount = 1 
 
 With ActiveDocument 
 Do 
 'Look for all paragraphs formatted with "Heading 4" style 
 If .Paragraphs(Index:=intCount).Style = "Heading 4" Then 
 .Paragraphs(Index:=intCount).Range.Select 
 
 'Insert a style separator if paragraph 
 'is formatted with a "Heading 4" style 
 Selection.InsertStyleSeparator 
 End If 
 intCount = intCount + 1 
 intParaCount = .Paragraphs.Count 
 Loop Until intCount = intParaCount 
 
 End With 
End Sub