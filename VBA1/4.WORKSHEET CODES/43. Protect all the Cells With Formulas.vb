Sub lockCellsWithFormulas()
With ActiveSheet
.Unprotect
.Cells.Locked = False
.Cells.SpecialCells(xlCellTypeFormulas).Locked = True
.Protect AllowDeletingRows:=True
End With
End Sub