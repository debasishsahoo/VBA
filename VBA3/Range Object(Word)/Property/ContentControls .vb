Dim objCC As ContentControl 
Dim objRange as Range 
 
Set objRange = ActiveDocument.Range(200, 200) 
Set objCC = objRange.ContentControls.Add(wdContentControlDropdownList) 
 
'List entries 
objCC.DropdownListEntries.Add "Cat" 
objCC.DropdownListEntries.Add "Dog" 
objCC.DropdownListEntries.Add "Horse" 
objCC.DropdownListEntries.Add "Monkey" 
objCC.DropdownListEntries.Add "Snake" 
objCC.DropdownListEntries.Add "Other"