Sub ChangeDocumentLayout() 
    With ActiveDocument.PageSetup 
        .LeftMargin = InchesToPoints(0.75) 
        .RightMargin = InchesToPoints(0.75) 
        .TopMargin = InchesToPoints(1.5) 
        .BottomMargin = InchesToPoints(1) 
    End With 
End Sub