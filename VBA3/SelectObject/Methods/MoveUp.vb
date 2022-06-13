Selection.MoveRight 
Selection.MoveUp Unit:=wdParagraph, Count:=2, Extend:=wdMove




MsgBox "Line " & Selection.Information(wdFirstCharacterLineNumber) 
Selection.MoveUp Unit:=wdLine, Count:=3, Extend:=wdMove 
MsgBox "Line " & Selection.Information(wdFirstCharacterLineNumber)