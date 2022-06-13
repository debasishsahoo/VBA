Set myRange = ActiveDocument.Range(0, 0) 
With myRange 
 .InsertBefore "Title" 
 .ParagraphFormat.Alignment = wdAlignParagraphCenter 
 .InsertParagraphAfter 
End With

ActiveDocument.Content.InsertParagraphAfter