ActiveDocument.Styles(wdStyleHeading1).Font.Bold = False

Set myRange = ActiveDocument.Paragraphs(2).Range 
If myRange.Font.Name = "Times New Roman" Then 
 myRange.Font.Name = "Arial" 
Else 
 myRange.Font.Name = "Times New Roman" 
End If