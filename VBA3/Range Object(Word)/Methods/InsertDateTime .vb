With ActiveDocument.Content 
 .Collapse Direction:=wdCollapseEnd 
 .InsertDateTime DateTimeFormat:="MM/dd/yy", _ 
 InsertAsField:=False 
End With

ActiveDocument.Sections(1).Footers(wdHeaderFooterPrimary).Range _ 
 .InsertDateTime DateTimeFormat:="MMMM dd, yyyy", _ 
 InsertAsField:=True
