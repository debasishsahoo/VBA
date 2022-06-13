For Each aFrame In ActiveDocument.Sections(1).Range.Frames 
 aFrame.TextWrap = True 
Next aFrame

Set myFrame = ActiveDocument.Frames.Add(Range:=Selection.Range)