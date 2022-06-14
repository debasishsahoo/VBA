'Modify cell (1,3) to "Hello World!"
With ActiveDocument.Tables(1)
        .Cell(1, 3).Range.Text = "Hello World!"
End With