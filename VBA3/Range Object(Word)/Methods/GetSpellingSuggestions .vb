If Selection.Range.GetSpellingSuggestions.Count = 0 Then 
 Msgbox "No suggestions." 
Else 
 Selection.Range.CheckSpelling 
End If