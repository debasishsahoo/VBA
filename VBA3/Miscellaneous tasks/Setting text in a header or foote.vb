Sub AddHeaderText() 
    With ActiveDocument.ActiveWindow.View 
        .SeekView = wdSeekCurrentPageHeader 
        Selection.HeaderFooter.Range.Text = "Header text" 
        .SeekView = wdSeekMainDocument 
    End With 
End Sub


Sub AddFooterText() 
    Dim rngFooter As Range 
    Set rngFooter = ActiveDocument.Sections(1) _ 
        .Footers(wdHeaderFooterPrimary).Range 
    With rngFooter 
        .Delete 
        .Fields.Add Range:=rngFooter, Type:=wdFieldFileName, Text:="\p" 
        .InsertAfter Text:=vbTab & vbTab 
        .Collapse Direction:=wdCollapseStart 
        .Fields.Add Range:=rngFooter, Type:=wdFieldAuthor 
    End With 
End Sub