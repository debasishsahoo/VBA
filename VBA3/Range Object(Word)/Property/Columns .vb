If ActiveDocument.Tables.Count >= 1 Then 
 MsgBox ActiveDocument.Tables(1).Columns.Count 
End If

If Selection.Information(wdWithInTable) = True Then 
 Selection.Columns.SetWidth ColumnWidth:=InchesToPoints(1), _ 
 RulerStyle:=wdAdjustProportional 
End If