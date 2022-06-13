Sub DeleteSelection() 
 Dim intResponse As Integer 
 
 intResponse = MsgBox("Are you sure you want to " & _ 
 "delete the contents of the document?", vbYesNo) 
 
 If intResponse = vbYes Then 
 ActiveDocument.Content.Select 
 Selection.Delete 
 End If 
End Sub