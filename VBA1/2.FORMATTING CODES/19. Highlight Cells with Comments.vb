Sub highlightCommentCells()
Selection.SpecialCells(xlCellTypeComments).Select
Selection.Style= "Note"
End Sub