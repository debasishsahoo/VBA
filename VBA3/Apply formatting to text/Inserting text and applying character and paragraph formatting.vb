Sub InsertFormatText() 
 Dim rngFormat As Range 
 Set rngFormat = ActiveDocument.Range(Start:=0, End:=0) 
 With rngFormat 
 .InsertAfter Text:="Title" 
 .InsertParagraphAfter 
 With .Font 
 .Name = "Tahoma" 
 .Size = 24 
 .Bold = True 
 End With 
 End With 
 With ActiveDocument.Paragraphs(1) 
 .Alignment = wdAlignParagraphCenter 
 .SpaceAfter = InchesToPoints(0.5) 
 End With 
End Sub