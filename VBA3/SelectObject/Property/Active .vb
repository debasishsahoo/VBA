Sub SplitWindow() 
 ActiveDocument.ActiveWindow.Split = True 
 If ActiveDocument.ActiveWindow.Panes(1).Selection _ 
 .Active = False Then 
 ActiveDocument.ActiveWindow.Panes(1).Activate 
 End If 
End Sub