Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdExtend


unitsMoved = Selection.MoveDown(Unit:=wdParagraph, _ 
 Count:=3, Extend:=wdMove) 
If unitsMoved = 3 Then Selection.Text = "Company"


MsgBox "Line " & Selection.Information(wdFirstCharacterLineNumber) 
Selection.MoveDown Unit:=wdLine, Count:=3, Extend:=wdMove 
MsgBox "Line " & Selection.Information(wdFirstCharacterLineNumber)