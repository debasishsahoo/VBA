With ActiveDocument.ActiveWindow.View 
 .Type = wdPrintView 
 .SeekView = wdSeekCurrentPageFooter 
End With 
Selection.HeaderFooter.PageNumbers.Add _ 
 PageNumberAlignment:=wdAlignPageNumberCenter