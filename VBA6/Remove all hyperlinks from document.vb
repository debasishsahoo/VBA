Sub RemoveHyperlinks()
  Dim oField As Field
  For Each oField In ActiveDocument.Fields
  If oField.Type = wdFieldHyperlink Then
    oField.Unlink
  End If
  Next
  Set oField = Nothing
End Sub