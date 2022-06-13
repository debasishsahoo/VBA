Select Case Selection.Range.Scripts(2).Language 
 Case msoScriptLanguageASP 
 MsgBox "Active Server Pages" 
 Case msoScriptLanguageVisualBasic 
 MsgBox "VBScript" 
 Case msoScriptLanguageJava 
 MsgBox "JavaScript" 
 Case msoScriptLanguageOther 
 MsgBox "Unknown type of script" 
End Select