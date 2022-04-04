Sub columnDifference()
Range("H7:H8,I7:I8").Select
Selection.ColumnDifferences(ActiveCell).Select
Selection.Style= "Bad"
End Sub