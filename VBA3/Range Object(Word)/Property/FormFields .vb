myType = ActiveDocument.Sections(2).Range.FormFields(1).Type 
Select Case myType 
 Case wdFieldFormTextInput 
 thetype = "TextBox" 
 Case wdFieldFormDropDown 
 thetype = "DropDown" 
 Case wdFieldFormCheckBox 
 thetype = "CheckBox" 
End Select