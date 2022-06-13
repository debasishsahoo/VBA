Sub ToggleParagraphSpace() 
 With Selection.Paragraphs(1) 
 If .SpaceBefore = 12 Then 
 .SpaceBefore = 0 
 Else 
 .SpaceBefore = 12 
 End If 
 End With 
End Sub