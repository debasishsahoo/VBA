Sub NewSampleDoc() 
    Dim docNew As Document 
    Set docNew = Documents.Add 
    With docNew 
        .Content.Font.Name = "Tahoma" 
        .SaveAs FileName:="Sample.doc" 
    End With 
End Sub