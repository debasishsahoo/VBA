char = Selection.StartOf(Unit:=wdLine, Extend:=wdMove)

Selection.Collapse Direction:=wdCollapseStart charmoved = Selection.StartOf(Unit:=wdLine, Extend:=wdExtend)

Selection.StartOf Unit:=wdParagraph, Extend:=wdMove

Set myRange = ActiveDocument.Sentences(2) 
myRange.StartOf Unit:=wdSentence, Extend:=wdMove 
myRange.Select