ActiveDocument.Words(1).Select 
MsgBox Selection.StartIsActive 
Selection.Flags = wdSelStartActive 
MsgBox Selection.StartIsActive

Selection.Flags = wdSelStartActive