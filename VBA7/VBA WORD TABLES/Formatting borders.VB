'Modify row 1 border, single, black and 100pt
 With ActiveDocument.Tables(1).Rows(3).Borders(wdBorderBottom)
        .LineStyle = wdLineStyleSingle
        .LineWidth = wdLineWidth100pt
        .Color = wdColorBlack
End With