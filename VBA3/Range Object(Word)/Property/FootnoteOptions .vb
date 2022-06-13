Sub SetFootnoteOptionsRange() 
 ActiveDocument.Sections(2).Range.FootnoteOptions _ 
 .NumberingRule = wdRestartSection 
End Sub