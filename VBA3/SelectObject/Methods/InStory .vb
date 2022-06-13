With ActiveDocument.ActiveWindow.View 
 .Type = wdPrintView 
 .SeekView = wdSeekCurrentPageHeader 
End With 
same = Selection.InStory(ActiveDocument.Paragraphs(1).Range) 
MsgBox same