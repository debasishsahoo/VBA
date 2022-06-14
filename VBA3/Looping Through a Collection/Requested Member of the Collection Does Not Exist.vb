Sub SelectTable() 
 ActiveDocument.Tables(1).Select 
End Sub

Sub SelectFirstTable() 
 If ActiveDocument.Tables.Count > 0 Then 
 ActiveDocument.Tables(1).Select 
 Else 
 MsgBox "Document doesn't contain a table" 
 End If 
End Sub

Sub DeleteAutoTextEntry() 
 Dim aceEntry As AutoCorrectEntry 
 For Each aceEntry In AutoCorrect.Entries 
 If aceEntry.Name = "acheive" Then aceEntry.Delete 
 Next aceEntry 
End Sub