MsgBox Selection.Font.Name

Set myFont = Selection.Font.Duplicate 
ActiveDocument.Paragraphs(1).Range.Font = myFont