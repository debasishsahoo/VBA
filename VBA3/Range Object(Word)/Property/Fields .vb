For Each aField in ActiveDocument.Fields 
 aField.Delete 
Next aField 
Set myRange = ActiveDocument.Sections(1).Footers _ 
 (wdHeaderFooterPrimary).Range 
For Each aField In myRange.Fields 
 aField.Delete 
Next aField