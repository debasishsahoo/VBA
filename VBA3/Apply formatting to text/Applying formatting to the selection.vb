Sub FormatSelection() 
 With Selection.Font 
 .Name = "Times New Roman" 
 .Size = 14 
 .AllCaps = True 
 End With 
 With Selection.ParagraphFormat 
 .LeftIndent = InchesToPoints(0.5) 
 .Space1 
 End With 
End Sub