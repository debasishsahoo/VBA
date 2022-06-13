Sub ActivateOrOpenDocument() 
    Dim doc As Document 
    Dim docFound As Boolean 
 
    For Each doc In Documents 
        If InStr(1, doc.Name, "sample.doc", 1) Then 
            doc.Activate 
            docFound = True 
            Exit For 
        Else 
            docFound = False 
        End If 
    Next doc 
 
    If docFound = False Then Documents.Open FileName:="Sample.doc" 
End Sub