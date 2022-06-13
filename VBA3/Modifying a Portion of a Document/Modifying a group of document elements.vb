Sub SetRangeForFirstTenCharacters() 
    Dim rngTenCharacters As Range 
    Set rngTenCharacters = ActiveDocument.Range(Start:=0, End:=10) 
End Sub


Sub SetRangeForFirstThreeWords() 
    Dim docActive As Document 
    Dim rngThreeWords As Range 
    Set docActive = ActiveDocument 
    Set rngThreeWords = docActive.Range(Start:=docActive.Words(1).Start, _ 
        End:=docActive.Words(3).End) 
End Sub

Sub SetParagraphRange() 
    Dim docActive As Document 
    Dim rngParagraphs As Range 
    Set docActive = ActiveDocument 
    Set rngParagraphs = docActive.Range(Start:=docActive.Paragraphs(2).Range.Start, _ 
        End:=docActive.Paragraphs(3).Range.End) 
End Sub