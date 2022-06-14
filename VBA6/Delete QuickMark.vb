Sub QuickMarkDelete()
  If ActiveDocument.Bookmarks.Exists("QuickMark") = True Then
    ActiveDocument.Bookmarks("QuickMark").Delete
  End If
End Sub