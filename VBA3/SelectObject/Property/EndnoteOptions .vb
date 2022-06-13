Sub SetEndnoteOptionsRange() 
 With Selection.EndnoteOptions 
 If .StartingNumber <> 1 Then 
 .StartingNumber = 1 
 End If 
 End With 
End Sub