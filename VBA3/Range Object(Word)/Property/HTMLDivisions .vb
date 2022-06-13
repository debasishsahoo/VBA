Sub FormatHTMLDivisions() 
 With ActiveDocument.Range.HTMLDivisions(1) 
 With .Borders(wdBorderLeft) 
 .Color = wdColorRed 
 .LineStyle = wdLineStyleSingle 
 End With 
 With .Borders(wdBorderRight) 
 .Color = wdColorRed 
 .LineStyle = wdLineStyleSingle 
 End With 
 With .HTMLDivisions(1) 
 .LeftIndent = InchesToPoints(1) 
 .RightIndent = InchesToPoints(1) 
 With .Borders(wdBorderTop) 
 .Color = wdColorBlue 
 .LineStyle = wdLineStyleDouble 
 End With 
 With .Borders(wdBorderBottom) 
 .Color = wdColorBlue 
 .LineStyle = wdLineStyleDouble 
 End With 
 With .HTMLDivisions(1) 
 .LeftIndent = InchesToPoints(1) 
 .RightIndent = InchesToPoints(1) 
 With .Borders(wdBorderLeft) 
 .LineStyle = wdLineStyleDot 
 End With 
 With .Borders(wdBorderRight) 
 .LineStyle = wdLineStyleDot 
 End With 
 With .Borders(wdBorderTop) 
 .LineStyle = wdLineStyleDot 
 End With 
 With .Borders(wdBorderBottom) 
 .LineStyle = wdLineStyleDot 
 End With 
 End With 
 End With 
 End With 
 
End Sub