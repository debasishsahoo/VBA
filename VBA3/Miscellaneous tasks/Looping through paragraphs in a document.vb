Sub LoopParagraphs() 
    Dim parCount As Paragraph 
    For Each parCount In ActiveDocument.Paragraphs 
        If parCount.SpaceBefore = 12 Then parCount.SpaceBefore = 6 
    Next parCount 
End Sub