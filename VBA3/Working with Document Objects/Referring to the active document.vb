Sub ActiveDocumentName() 
    If Documents.Count >= 1 Then 
        MsgBox ActiveDocument.Name 
    Else 
        MsgBox "No documents are open" 
    End If 
End Sub