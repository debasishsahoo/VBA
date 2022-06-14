Sub Save_Documnent()
    If ActiveDocument.Saved = False Then
        ActiveDocument.Save
    End If
End Sub