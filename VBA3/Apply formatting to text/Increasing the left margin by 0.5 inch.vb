Sub FormatMargins() 
 With ActiveDocument.PageSetup 
 .LeftMargin = .LeftMargin + InchesToPoints(0.5) 
 .RightMargin = .RightMargin + InchesToPoints(0.5) 
 End With 
End Sub