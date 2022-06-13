Dim countOfUpdates As Integer 
 
countOfUpdates = ActiveDocument.Paragraphs(1).Range.Updates.Count 
 
MsgBox "The number of updates is " & countOfUpdates