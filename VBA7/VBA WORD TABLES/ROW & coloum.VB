With ActiveDocument.Tables(1)
        ' Modify height of row 1
        .Rows(1).Height = CentimetersToPoints(0.65)
        ' Modify width of column 1
        .Columns(1).Width = CentimetersToPoints(2.54)
        'Merge cell (1,3) with cell (1,4) - cells must be next to each other
        .Cell(1, 3).Merge .Cell(1, 4)
End With