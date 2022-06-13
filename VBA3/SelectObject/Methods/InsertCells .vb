If Selection.Cells.Count >= 1 Then 
 Selection.InsertCells ShiftCells:=wdInsertCellsShiftRight 
 For Each aBorder In Selection.Borders 
 aBorder.LineStyle = wdLineStyleSingle 
 aBorder.ColorIndex = wdRed 
 Next aBorder 
End If