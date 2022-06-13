MsgBox "The selection is on page " & _ 
 Selection.Information(wdActiveEndPageNumber) & " of page " _ 
 & Selection.Information(wdNumberOfPagesInDocument)
 
 If Selection.Information(wdWithInTable) Then _ 
 Selection.Tables(1).Select

 Selection.Collapse Direction:=wdCollapseStart 
MsgBox "The insertion point is in section " & _ 
 Selection.Information(wdActiveEndSectionNumber)