char = Selection.EndOf(Unit:=wdWord, Extend:=wdMove)

charmoved = Selection.EndOf(Unit:=wdParagraph, Extend:=wdExtend) 
If charmoved = 0 Then MsgBox "Selection unchanged"

Set myRange = Selection.Characters(1) 
myRange.EndOf Unit:=wdWord, Extend:=wdMove

Set myRange = ActiveDocument.Range(0, 0) 
Set myTable = ActiveDocument.Tables.Add(Range:=myRange, _ 
 NumRows:=5, NumColumns:=3) 
myTable.Cell(2, 1).Select 
Selection.EndOf Unit:=wdColumn, Extend:=wdExtend