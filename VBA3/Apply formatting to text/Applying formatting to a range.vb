Sub FormatRange() 
 Dim rngFormat As Range 
 Set rngFormat = ActiveDocument.Range( _ 
 Start:=ActiveDocument.Paragraphs(1).Range.Start, _ 
 End:=ActiveDocument.Paragraphs(3).Range.End) 
 With rngFormat 
 .Font.Name = "Arial" 
 .ParagraphFormat.Alignment = wdAlignParagraphJustify 
 End With 
End Sub