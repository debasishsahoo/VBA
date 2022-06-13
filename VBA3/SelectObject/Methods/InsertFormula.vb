Set MyRange = ActiveDocument.Range(0, 0) 
Set myTable = ActiveDocument.Tables.Add(MyRange, 3, 3) 
With myTable 
 .Cell(1, 1).Range.InsertAfter "100" 
 .Cell(2, 1).Range.InsertAfter "50" 
 .Cell(3, 1).Select 
End With 
Selection.InsertFormula Formula:="=Average(Above)"

Selection.Collapse Direction:=wdCollapseStart 
Selection.InsertFormula Formula:= "=GrossSales-45,000.00", _ 
 NumberFormat:="$#,##0.00"