If Selection.Information(wdWithInTable) = True Then 
 Selection.Columns.SetWidth ColumnWidth:=InchesToPoints(1), _ 
 RulerStyle:=wdAdjustProportional 
End If