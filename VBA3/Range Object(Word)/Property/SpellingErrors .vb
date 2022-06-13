Set myErrors = ActiveDocument.Paragraphs(3).Range.SpellingErrors 
If myErrors.Count = 0 Then 
 Msgbox "No spelling errors found." 
Else 
 For Each myErr in myErrors 
 Msgbox myErr.Text 
 Next 
End If