Set myRange = ActiveDocument.Words(1) 
myRange.Expand Unit:=wdParagraph

With Selection 
 .Characters(1).Case = wdTitleSentence 
 .Expand Unit:=wdSentence 
End With