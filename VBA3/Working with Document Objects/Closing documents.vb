Sub CloseDocument()
    Documents("Sales.doc").Close SaveChanges:=wdSaveChanges 
End Sub
Sub CloseAllDocuments() 
    Documents.Close SaveChanges:=wdDoNotSaveChanges
End Sub
Sub PromptToSaveAndClose()
    Dim doc As Document
    For Each doc In Documents 
        doc.Close SaveChanges:=wdPromptToSaveChanges
    Next
End Sub