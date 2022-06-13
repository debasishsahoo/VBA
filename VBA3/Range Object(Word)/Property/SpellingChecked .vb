Set myRange = ActiveDocument.Sections(1).Range 
isChecked = myRange.SpellingChecked 
If isChecked = False Then 
 myRange.CheckSpelling 
Else 
 MsgBox "Spelling has already been checked in the range." 
End If