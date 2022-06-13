Sub SetEndnoteOptionsRange() 
 With ActiveDocument.Sections(2).Range.EndnoteOptions 
 If .StartingNumber <> 1 Then 
 .StartingNumber = 1 
 End If 
 End With 
End Sub