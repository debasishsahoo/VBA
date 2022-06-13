Sub SetParagraphRange() 
 Dim rngParagraph As Range 
 Set rngParagraph = ActiveDocument.Paragraphs(1).Range 
End Sub

Sub FormatRange() 
 ActiveDocument.Paragraphs(2).Range.Select 
 Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter 
End Sub

Sub FormatFirstParagraph() 
 Dim rngParagraph As Range 
 Set rngParagraph = ActiveDocument.Paragraphs(1).Range 
 With rngParagraph 
 .Bold = True 
 .ParagraphFormat.Alignment = wdAlignParagraphCenter 
 With .Font 
 .Name = "Stencil" 
 .Size = 15 
 End With 
 End With 
End Sub