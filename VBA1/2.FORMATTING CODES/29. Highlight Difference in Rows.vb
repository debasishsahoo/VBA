Sub rowDifference()
Range("H7:H8,I7:I8").Select
Selection.RowDifferences(ActiveCell).Select
Selection.Style= "Bad"
End Sub