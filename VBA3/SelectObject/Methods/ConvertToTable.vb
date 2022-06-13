With Selection 
 .Collapse 
 .InsertBefore "one, two, three" 
 .InsertParagraphAfter 
 .InsertAfter "one, two, three" 
 .InsertParagraphAfter 
End With 
Set myTable = Selection.ConvertToTable( _ 
 Separator:=wdSeparateByCommas, _ 
 Format:=wdTableFormatList8)