Set myRange = Documents("MyDoc.doc").Content 
With myRange.ParagraphFormat 
 .Space2 
 .TabStops.Add Position:=InchesToPoints(.25) 
End With