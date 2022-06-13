Sub CopyWord() 
    Selection.Words(1).Copy 
End Sub

Sub CopyParagraph() 
    ActiveDocument.Paragraphs(1).Range.Copy 
End Sub

Sub ChangeCase() 
    ActiveDocument.Words(1).Case = wdUpperCase 
End Sub

Sub ChangeSectionMargin() 
    Selection.Sections(1).PageSetup.BottomMargin = InchesToPoints(0.5) 
End Sub

Sub DoubleSpaceDocument() 
    ActiveDocument.Content.ParagraphFormat.Space2 
End Sub